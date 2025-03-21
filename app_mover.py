import os
import sys
import shutil
import winreg
import ctypes
import subprocess
from typing import List, Dict, Tuple

# Set console to support ANSI escape codes for better formatting
os.system('')

# ANSI color codes
COLORS = {
    'HEADER': '\033[95m',
    'BLUE': '\033[94m',
    'GREEN': '\033[92m',
    'YELLOW': '\033[93m',
    'RED': '\033[91m',
    'ENDC': '\033[0m',
    'BOLD': '\033[1m',
    'UNDERLINE': '\033[4m'
}

def clear_screen():
    """Clear the console screen."""
    os.system('cls' if os.name == 'nt' else 'clear')

def print_header():
    """Print the application header."""
    print(f"{COLORS['HEADER']}{COLORS['BOLD']}===================================={COLORS['ENDC']}")
    print(f"{COLORS['HEADER']}{COLORS['BOLD']}  APPLICATION FILE MOVER UTILITY  {COLORS['ENDC']}")
    print(f"{COLORS['HEADER']}{COLORS['BOLD']}===================================={COLORS['ENDC']}")
    print()

def is_admin():
    """Check if the script is running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False

def get_installed_applications() -> List[Dict]:
    """Get a list of installed applications from the Windows registry."""
    applications = []
    registry_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
    ]
    
    for reg_root, reg_path in registry_paths:
        try:
            registry_key = winreg.OpenKey(reg_root, reg_path)
            for i in range(winreg.QueryInfoKey(registry_key)[0]):
                try:
                    subkey_name = winreg.EnumKey(registry_key, i)
                    subkey = winreg.OpenKey(registry_key, subkey_name)
                    
                    try:
                        display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                        install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                        
                        # Skip entries without a display name or install location
                        if not display_name or not install_location:
                            continue
                            
                        # Skip Windows components and updates
                        if "KB" in display_name or "Microsoft Windows" in display_name:
                            continue
                            
                        # Get additional information if available
                        try:
                            publisher = winreg.QueryValueEx(subkey, "Publisher")[0]
                        except:
                            publisher = "Unknown"
                            
                        try:
                            uninstall_string = winreg.QueryValueEx(subkey, "UninstallString")[0]
                        except:
                            uninstall_string = ""
                            
                        # Add to our list if not already present
                        app_info = {
                            "name": display_name,
                            "publisher": publisher,
                            "install_location": install_location,
                            "uninstall_string": uninstall_string,
                            "registry_key": f"{reg_path}\\{subkey_name}",
                            "registry_root": reg_root
                        }
                        
                        # Check if this app is already in our list (avoid duplicates)
                        if not any(app["name"] == display_name for app in applications):
                            applications.append(app_info)
                            
                    except (WindowsError, FileNotFoundError):
                        pass
                    finally:
                        winreg.CloseKey(subkey)
                except (WindowsError, FileNotFoundError):
                    continue
            winreg.CloseKey(registry_key)
        except (WindowsError, FileNotFoundError):
            continue
    
    # Sort applications alphabetically by name
    return sorted(applications, key=lambda x: x["name"].lower())

def display_applications(applications: List[Dict]):
    """Display the list of applications with numbers."""
    print(f"{COLORS['BOLD']}Available Applications:{COLORS['ENDC']}")
    print(f"{COLORS['UNDERLINE']}{'No.':<4} {'Application Name':<50} {'Publisher':<30} {'Install Location'}{COLORS['ENDC']}")
    
    for i, app in enumerate(applications, 1):
        print(f"{i:<4} {app['name'][:48]:<50} {app['publisher'][:28]:<30} {app['install_location']}")

def move_application_files(app: Dict, destination_path: str) -> bool:
    """Move application files to the specified destination and update settings."""
    source_path = app["install_location"]
    app_name = app["name"]
    
    # Create destination directory if it doesn't exist
    os.makedirs(destination_path, exist_ok=True)
    
    print(f"\n{COLORS['YELLOW']}Moving files for {app_name} from {source_path} to {destination_path}...{COLORS['ENDC']}")
    
    try:
        # Check if source directory exists and is not empty
        if not os.path.exists(source_path) or not os.listdir(source_path):
            print(f"{COLORS['RED']}Error: Source directory is empty or does not exist.{COLORS['ENDC']}")
            return False
        
        # Get list of files and directories in the source
        items = os.listdir(source_path)
        
        # Move each item
        for item in items:
            source_item = os.path.join(source_path, item)
            dest_item = os.path.join(destination_path, item)
            
            if os.path.isdir(source_item):
                if os.path.exists(dest_item):
                    shutil.rmtree(dest_item)
                shutil.copytree(source_item, dest_item)
                shutil.rmtree(source_item)
                print(f"Moved directory: {item}")
            else:
                if os.path.exists(dest_item):
                    os.remove(dest_item)
                shutil.copy2(source_item, dest_item)
                os.remove(source_item)
                print(f"Moved file: {item}")
        
        # Update registry to point to the new location
        try:
            reg_root = app["registry_root"]
            reg_key_path = app["registry_key"]
            
            key = winreg.OpenKey(reg_root, reg_key_path, 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, destination_path)
            winreg.CloseKey(key)
            
            print(f"{COLORS['GREEN']}Successfully updated registry settings for {app_name}.{COLORS['ENDC']}")
        except Exception as e:
            print(f"{COLORS['RED']}Warning: Could not update registry settings: {str(e)}{COLORS['ENDC']}")
            print(f"{COLORS['YELLOW']}You may need to manually update the application settings.{COLORS['ENDC']}")
        
        # Create a symbolic link from the original location to the new location
        try:
            # Remove the original directory if it still exists
            if os.path.exists(source_path):
                os.rmdir(source_path)
                
            # Create a symbolic link
            os.symlink(destination_path, source_path, target_is_directory=True)
            print(f"{COLORS['GREEN']}Created symbolic link from {source_path} to {destination_path}.{COLORS['ENDC']}")
        except Exception as e:
            print(f"{COLORS['YELLOW']}Warning: Could not create symbolic link: {str(e)}{COLORS['ENDC']}")
            print(f"{COLORS['YELLOW']}Some applications may still look for files in the original location.{COLORS['ENDC']}")
        
        return True
    except Exception as e:
        print(f"{COLORS['RED']}Error moving files: {str(e)}{COLORS['ENDC']}")
        return False

def main():
    # Check for admin privileges
    if not is_admin():
        print(f"{COLORS['RED']}This script requires administrator privileges to modify application settings.{COLORS['ENDC']}")
        print(f"{COLORS['YELLOW']}Please run this script as administrator.{COLORS['ENDC']}")
        
        # Attempt to restart with admin privileges
        if os.name == 'nt':
            print("Attempting to restart with administrator privileges...")
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        sys.exit(1)
    
    clear_screen()
    print_header()
    
    print(f"{COLORS['YELLOW']}Scanning for installed applications...{COLORS['ENDC']}")
    applications = get_installed_applications()
    
    if not applications:
        print(f"{COLORS['RED']}No applications with valid install locations found.{COLORS['ENDC']}")
        input("Press Enter to exit...")
        return
    
    while True:
        clear_screen()
        print_header()
        display_applications(applications)
        
        print(f"\n{COLORS['BOLD']}Enter the number of the application to move, or 'q' to quit:{COLORS['ENDC']}")
        choice = input("> ").strip().lower()
        
        if choice == 'q':
            break
        
        try:
            app_index = int(choice) - 1
            if 0 <= app_index < len(applications):
                selected_app = applications[app_index]
                
                print(f"\n{COLORS['GREEN']}Selected: {selected_app['name']}{COLORS['ENDC']}")
                print(f"Current location: {selected_app['install_location']}")
                
                print(f"\n{COLORS['BOLD']}Enter the destination path (or 'c' to cancel):{COLORS['ENDC']}")
                dest_path = input("> ").strip()
                
                if dest_path.lower() == 'c':
                    continue
                
                # Validate destination path
                if not os.path.isabs(dest_path):
                    print(f"{COLORS['RED']}Error: Please enter an absolute path.{COLORS['ENDC']}")
                    input("Press Enter to continue...")
                    continue
                
                # Confirm the operation
                print(f"\n{COLORS['YELLOW']}Warning: This will move all files from:{COLORS['ENDC']}")
                print(f"  {selected_app['install_location']}")
                print(f"{COLORS['YELLOW']}to:{COLORS['ENDC']}")
                print(f"  {dest_path}")
                print(f"\n{COLORS['BOLD']}Are you sure you want to continue? (y/n){COLORS['ENDC']}")
                
                confirm = input("> ").strip().lower()
                if confirm != 'y':
                    continue
                
                # Perform the move operation
                success = move_application_files(selected_app, dest_path)
                
                if success:
                    print(f"\n{COLORS['GREEN']}Successfully moved {selected_app['name']} to {dest_path}.{COLORS['ENDC']}")
                    # Update the application's install location in our list
                    applications[app_index]['install_location'] = dest_path
                
                input("\nPress Enter to continue...")
            else:
                print(f"{COLORS['RED']}Invalid selection. Please enter a number between 1 and {len(applications)}.{COLORS['ENDC']}")
                input("Press Enter to continue...")
        except ValueError:
            print(f"{COLORS['RED']}Invalid input. Please enter a number.{COLORS['ENDC']}")
            input("Press Enter to continue...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
    except Exception as e:
        print(f"\n{COLORS['RED']}An unexpected error occurred: {str(e)}{COLORS['ENDC']}")
        input("Press Enter to exit...")