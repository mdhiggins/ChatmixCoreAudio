import usb.core
import usb.util
import hid
import pythoncom
import argparse
import logging
import os
import sys
import threading
import time
import json

from enum import Enum
from psutil import NoSuchProcess
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume


class CoreMix:
    def __init__(self, vendor_id, product_id, interface_number, endpoint_address, voice_apps, exclude_apps):
        """
        Initialize the CoreMix class with the necessary parameters.
        """
        # Setup logging
        self.log = logging.getLogger(__name__)

        # Initialize the volume levels
        self.voice_level = 100
        self.system_level = 100

        # Parameters passed directly to the constructor
        self.vendor_id = vendor_id
        self.product_id = product_id
        self.interface_number = interface_number
        self.endpoint_address = endpoint_address
        self.voice_apps = set(voice_apps)
        self.exclude_apps = set(exclude_apps)

        # Tracking of sessions
        self.known_sessions = set()
        self.voice_ids = set()
        self.exclude_ids = set()

        # Set the initial values for USB device and session
        self.dev = None
        self.interface = None
        self.endpoint = None

        # Set up exit event
        self.exit_event = threading.Event()

        # Automatically find the USB device during initialization
        self.find_usb_device()


    def find_usb_device(self):
        """
        Find and configure the USB device.
        """
        self.dev = usb.core.find(idVendor=self.vendor_id, idProduct=self.product_id)
        if self.dev is None:
            raise ValueError('Device not found')

        self.dev.set_configuration()

        for cfg in self.dev:
            for intf in cfg:
                if intf.bInterfaceNumber == self.interface_number:
                    self.interface = intf
                    break
            if self.interface is not None:
                break

        if self.interface is None:
            raise ValueError(f"Couldn't find interface {self.interface_number}")

        self.endpoint = usb.util.find_descriptor(
            self.interface,
            custom_match=lambda e: \
                usb.util.endpoint_direction(e.bEndpointAddress) == usb.util.ENDPOINT_IN and \
                e.bEndpointAddress == self.endpoint_address
        )

        if self.endpoint is None:
            raise ValueError(f"Couldn't find endpoint {self.endpoint_address} on interface {self.interface_number}")

    def usb_reader(self):
        """
        Thread function for reading data from the USB device.
        """
        self.log.info("Starting USB chatmix dial position monitoring thread")
        pythoncom.CoInitialize()
        try:
            while not self.exit_event.is_set():
                try:
                    data = self.dev.read(self.endpoint.bEndpointAddress, self.endpoint.wMaxPacketSize, timeout=0)
                    if data is not None:
                        voice_level = data[2]
                        system_level = data[1]

                        if voice_level == 100 or system_level == 100:
                            self.log.info(f'Voice Level: {voice_level}, System Level: {system_level}')
                            self.voice_level = voice_level
                            self.system_level = system_level
                            self.set_volume_levels(self.voice_level, self.system_level)
                        else:
                            self.log.debug('Voice or system level not at 100, ignoring this from the headset')
                except usb.core.USBError as e:
                    if e.errno == 110:  # Timeout error, just continue
                        continue
                    elif e.errno == 19:  # Device disconnected
                        self.log.info("USB device disconnected.")
                        break
                    else:
                        self.log.exception(f"USB error: {e}")
                except KeyboardInterrupt:
                    raise
        except KeyboardInterrupt:
            raise
        except Exception as e:
            self.log.exception(f"Unexpected error in USB thread: {e}")
        finally:
            usb.util.dispose_resources(self.dev)

    def monitor_new_sessions(self):
        """
        Polling function for monitoring new audio sessions.
        """
        self.log.info("Starting audio session monitoring thread")
        pythoncom.CoInitialize()
        self.known_sessions = set((session.Process.name(), session.ProcessId) for session in AudioUtilities.GetAllSessions() if session.Process)
        while not self.exit_event.is_set():
            try:
                current_sessions = set((session.Process.name(), session.ProcessId) for session in AudioUtilities.GetAllSessions() if session.Process)
                new_sessions = current_sessions - self.known_sessions
                if new_sessions:
                    print(self.voice_apps)
                    for app_name, app_id in new_sessions:
                        print(app_name)
                        if app_name in self.voice_apps:
                            self.voice_ids.add(app_id)
                            self.log.info(f"New voice audio session: {app_name} {app_id}")
                        elif app_name in self.exclude_apps:
                            self.exclude_ids.add(app_id)
                            self.log.debug(f"Ignoring new audio session: {app_name} {app_id}")
                        else:
                            self.log.info(f"New general audio session: {app_name} {app_id}")
                    self.set_volume_levels(self.voice_level, self.system_level)

                self.known_sessions = current_sessions
                time.sleep(1)
            except NoSuchProcess:
                continue

    def set_volume_levels(self, voice_level, system_level):
        """
        Set the volume levels for the voice and system applications.
        """
        sessions = AudioUtilities.GetAllSessions()

        for session in sessions:
            if session.Process:
                volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                app_name = session.Process.name()
                app_id = session.ProcessId

                if app_id in self.voice_ids:
                    self.log.debug(f"Setting volume for {app_name} to {voice_level}")
                    volume.SetMasterVolume(voice_level / 100, None)
                elif app_id in self.exclude_ids:
                    self.log.debug(f"Ignoring volume for {app_name}")
                else:
                    self.log.debug(f"Setting volume for {app_name} {app_id} to {system_level}")
                    volume.SetMasterVolume(system_level / 100, None)

    def run(self):
        """
        Start the threads for USB reading and audio session monitoring.
        """
        # Start USB reading thread
        usb_thread = threading.Thread(target=self.usb_reader)
        usb_thread.start()

        # Start audio sessions monitoring thread 
        monitor_thread = threading.Thread(target=self.monitor_new_sessions)
        monitor_thread.start()

        try:
            while usb_thread.is_alive() or monitor_thread.is_alive():
                usb_thread.join(timeout=0.1)
                monitor_thread.join(timeout=0.1)
        except KeyboardInterrupt:
            self.log.info("Exiting...")
            self.exit_event.set()
            usb_thread.join()
        except Exception as e:
            self.log.exception(f"Unexpected error: {e}")
        finally:
            usb.util.dispose_resources(self.dev)
            if usb_thread.is_alive():
                usb_thread.join()

            sys.exit(0)

    def stop(self):
        """
        Stop the USB reading and session monitoring threads gracefully.
        """
        self.log.info("Stopping the application...")
        self.exit_event.set()  # Signal the threads to exit
        # Wait for the threads to finish
        usb_thread = threading.Thread(target=self.usb_reader)
        monitor_thread = threading.Thread(target=self.monitor_new_sessions)
        usb_thread.join()
        monitor_thread.join()
        self.log.info("Application stopped gracefully.")


def parse_arguments():
    parser = argparse.ArgumentParser(description="USB Volume Control")
    parser.add_argument("--config", type=str, help="Path to the configuration file (optional)")
    parser.add_argument("--debug", action="store_true", help="Enable debug level logging")
    parser.add_argument("--vendor_id", type=int, help="Vendor ID of the USB device")
    parser.add_argument("--product_id", type=int, help="Product ID of the USB device")
    parser.add_argument("--interface_number", type=int, help="Interface number of the USB device")
    parser.add_argument("--endpoint_address", type=int, help="Endpoint address of the USB device")
    parser.add_argument("--voice_apps", nargs="*", help="List of voice application names")
    parser.add_argument("--exclude_apps", nargs="*", help="List of apps to exclude from volume control")

    return parser.parse_args()


def load_config(file_path, logger):
    if not os.path.isabs(file_path):
        file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_path)

    config = {}
    try:
        with open(file_path, "r") as file:
            config = json.load(file)

        def convert_hex(value):
            if isinstance(value, str):
                if value.startswith("0x"):
                    try:
                        return int(value, 16)
                    except ValueError:
                        logger.warning(f"Invalid hexadecimal value: {value}")
                else:
                    return int(value)
            return value

        if "vendor_id" in config:
            config["vendor_id"] = convert_hex(config["vendor_id"])
        if "product_id" in config:
            config["product_id"] = convert_hex(config["product_id"])
        if "endpoint_address" in config:
            config["endpoint_address"] = convert_hex(config["endpoint_address"])

    except FileNotFoundError:
        logger.error(f"Configuration file '{file_path}' not found!")
    except json.JSONDecodeError as e:
        logger.error(f"Error parsing configuration file '{file_path}': {e}")
    except Exception as e:
        logger.exception(f"Error reading configuration file: {e}")
    
    return config


if __name__ == "__main__":
    args = parse_arguments()

    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO if not args.debug else logging.DEBUG)
    stdout_handler = logging.StreamHandler()
    stdout_handler.setLevel(logging.DEBUG)
    stdout_handler.setFormatter(logging.Formatter('%(levelname)8s | %(message)s'))
    logger.addHandler(stdout_handler)

    # Load configuration if a config file is provided
    config = {}
    if args.config:
        config = load_config(args.config, logger)

    vendor_id = args.vendor_id or config.get("vendor_id", 0x1038)
    product_id = args.product_id or config.get("product_id", 0x2202)
    interface_number = args.interface_number or config.get("interface_number", 5)
    endpoint_address = args.endpoint_address or config.get("endpoint_address", 0x86)
    voice_apps = set(args.voice_apps or config.get("voice_apps", []))
    exclude_apps = set(args.exclude_apps or config.get("exclude_apps", []))

    coremix = CoreMix(vendor_id, product_id, interface_number, endpoint_address, voice_apps, exclude_apps)

    # Run the application
    coremix.run()
