import usb.core
import usb.util
import pythoncom
import argparse
import logging
import os
import sys
import threading
import time
import json

from psutil import NoSuchProcess
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

# Setup logging
log = logging.getLogger(__name__)
log.setLevel(logging.INFO)
stdout_handler = logging.StreamHandler()
stdout_handler.setLevel(logging.DEBUG)
stdout_handler.setFormatter(logging.Formatter('%(levelname)8s | %(message)s'))
log.addHandler(stdout_handler)

# Default configuration file
DEFAULT_CONFIG_FILE = "config.json"

# Read the configuration from a JSON file
def load_config(file_path=DEFAULT_CONFIG_FILE):
    if not os.path.isabs(file_path):
        file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_path)

    config = {}
    try:
        with open(file_path, "r") as file:
            config = json.load(file)

        # Check and convert hexadecimal string values to integers
        def convert_hex(value):
            if isinstance(value, str):
                # Check if the string starts with '0x' and try to convert
                if value.startswith("0x"):
                    try:
                        return int(value, 16)  # Convert from hex string to integer
                    except ValueError:
                        log.warning(f"Invalid hexadecimal value: {value}")
                else:
                    return int(value)  # Return the int value of the string if it's not in hex format
            return value  # Return the value as is if it's already an int or something else

        # Apply the conversion function to the relevant config values
        if "vendor_id" in config:
            config["vendor_id"] = convert_hex(config["vendor_id"])
        if "product_id" in config:
            config["product_id"] = convert_hex(config["product_id"])
        if "endpoint_address" in config:
            config["endpoint_address"] = convert_hex(config["endpoint_address"])

    except FileNotFoundError:
        log.error(f"Configuration file '{file_path}' not found!")
    except json.JSONDecodeError as e:
        log.error(f"Error parsing configuration file '{file_path}': {e}")
    except Exception as e:
        log.exception(f"Error reading configuration file: {e}")
    
    return config


# Parse command-line arguments
def parse_arguments():
    parser = argparse.ArgumentParser(description="USB Volume Control")

    # Add argument for the configuration file path
    parser.add_argument("--config", type=str, default=DEFAULT_CONFIG_FILE, help="Path to the configuration file (default: 'config.json')")
    parser.add_argument("--debug", action="store_true", help="Enable debug level logging (default: False)")

    return parser.parse_args()

# Main function to run the script
def main():
    # Parse arguments
    args = parse_arguments()

    # Set logging level
    log.setLevel(logging.DEBUG if args.debug else logging.INFO)

    # Load configuration from JSON file (using the file path from the argument)
    config = load_config(args.config)


    # Get configuration values from the JSON file (fall back to defaults if not specified)
    vendor_id = config.get("vendor_id", 0x1038)
    product_id = config.get("product_id", 0x2202)
    interface_number = config.get("interface_number", 5)
    endpoint_address = config.get("endpoint_address", 0x86)
    voice_apps = set(config.get("voice_apps", []))
    exclude_apps = set(config.get("exclude_apps", []))

    # Find the USB device
    dev = usb.core.find(idVendor=vendor_id, idProduct=product_id)

    if dev is None:
        raise ValueError('Device not found')

    # Set the active configuration
    dev.set_configuration()

    # Find the specific interface by its number
    interface_number = interface_number
    endpoint_address = endpoint_address
    interface = None
    for cfg in dev:
        for intf in cfg:
            if intf.bInterfaceNumber == interface_number:
                interface = intf
                break
        if interface is not None:
            break

    if interface is None:
        raise ValueError(f"Couldn't find interface {interface_number}")

    # Find the specific endpoint on the interface
    endpoint = usb.util.find_descriptor(
        interface,
        custom_match=lambda e: \
            usb.util.endpoint_direction(e.bEndpointAddress) == usb.util.ENDPOINT_IN and \
            e.bEndpointAddress == endpoint_address
    )

    if endpoint is None:
        raise ValueError(f"Couldn't find endpoint {endpoint_address} on interface {interface_number}")

    # Event to signal thread termination
    exit_event = threading.Event()

    # Cache for voice and system levels
    volume_cache = {
        "voice_level": 100,
        "system_level": 100
    }

    # Method for setting the levels
    def set_volume_levels(voice_level, system_level):
        # Get all active audio sessions
        sessions = AudioUtilities.GetAllSessions()

        # Iterate through each session
        for session in sessions:
            if session.Process:
                volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                app_name = session.Process.name()
                app_id = session.ProcessId

                # Check if the application is in the voice list
                if app_name in voice_apps:
                    # Apply voice volume level
                    log.debug(f"Setting volume for {app_name} to {voice_level}")
                    log.debug(f"Setting volume for {app_name} {app_id} to {voice_level}")

                    volume.SetMasterVolume(voice_level / 100, None)
                elif app_name in exclude_apps:
                    log.debug(f"Ignoring volume for {app_name}")
                    log.debug(f"Ignoring volume for {app_name} {app_id}")
                else:
                    # Apply system volume level
                    log.debug(f"Setting volume for {app_name} {app_id} to {system_level}")
                    volume.SetMasterVolume(system_level / 100, None)

    # Thread function for USB reading
    def usb_reader():
        log.info("Starting USB chatmix dial position monitoring thread")
        pythoncom.CoInitialize()
        try:
            while not exit_event.is_set():
                try:
                    data = dev.read(endpoint.bEndpointAddress, endpoint.wMaxPacketSize, timeout=0)  # Non-blocking read
                    if data is not None:
                        # Extract voice and system levels from the received data
                        voice_level  = data[2]  
                        system_level = data[1]

                        # Data should only be sent if one of the levels is 100, because for some reason sometimes the headset sends an incorrect amount
                        if voice_level == 100 or system_level == 100:
                            log.info(f'Voice Level: {voice_level}, System Level: {system_level}')
                            
                            # Update the cache with the new levels
                            volume_cache["voice_level"] = voice_level
                            volume_cache["system_level"] = system_level

                            # Set the volume levels
                            set_volume_levels(volume_cache["voice_level"], volume_cache["system_level"])
                        else:
                            log.debug('Voice or system level not at 100, ignoring this from the headset')
                except usb.core.USBError as e:
                    if e.errno == 110:  # errno 110 is a timeout error
                        continue  # Non-blocking read, continue to check for KeyboardInterrupt
                    elif e.errno == 19:  # errno 19 is no such device error
                        log.info("USB device disconnected.")
                        break  # Exit the loop on device disconnection
                    else:
                        log.exception(f"USB error: {e}")
                except KeyboardInterrupt:
                    raise
        except KeyboardInterrupt:
            raise
        except Exception as e:
            log.exception(f"Unexpected error in USB thread: {e}")
        finally:
            usb.util.dispose_resources(dev)

    # Polling function for monitoring new audio sessions
    def monitor_new_sessions():
        log.info("Starting audio session monitoring thread")
        pythoncom.CoInitialize()
        known_sessions = set(session.Process.name() for session in AudioUtilities.GetAllSessions() if session.Process)
        while not exit_event.is_set():
            # Get the current active audio sessions
            try:
                # Extract the names of the processes in active sessions
                current_sessions = set(session.Process.name() for session in AudioUtilities.GetAllSessions() if session.Process)

                # Identify new sessions that were not present before
                new_sessions = current_sessions - known_sessions
                for app_name in new_sessions:
                    log.info(f"New audio session created: {app_name}")

                    # Apply the volume levels (voice_level, system_level)
                    time.sleep(1)
                    set_volume_levels(volume_cache["voice_level"], volume_cache["system_level"])

                # Update the known sessions
                known_sessions = current_sessions

                # Sleep to reduce CPU usage
                time.sleep(1)
            except NoSuchProcess:
                continue

    # Start USB reading thread
    usb_thread = threading.Thread(target=usb_reader)
    usb_thread.start()

    # Start monitoring thread for audio sessions using polling
    monitor_thread = threading.Thread(target=monitor_new_sessions)
    monitor_thread.start()

    try:
        # Wait for KeyboardInterrupt
        while usb_thread.is_alive() or monitor_thread.is_alive():
            usb_thread.join(timeout=0.1)  # Check thread status with timeout
            monitor_thread.join(timeout=0.1)  # Check thread status with timeout
    except KeyboardInterrupt:
        log.info("Exiting...")
        # Set exit event to stop the threads
        exit_event.set()
        usb_thread.join()
    except Exception as e:
        log.exception(f"Unexpected error: {e}")
    finally:
        # Ensure the device is properly released
        usb.util.dispose_resources(dev)

        # Wait for threads to complete if they're still running
        if usb_thread.is_alive():
            usb_thread.join()

        sys.exit(0)


if __name__ == "__main__":
    main()
