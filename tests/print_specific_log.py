import os
import sys

def print_log_content(log_file_path):
    """Print the content of the specified log file."""
    try:
        # First try UTF-8
        try:
            with open(log_file_path, 'r', encoding='utf-8') as log_file:
                content = log_file.read()
                print(content)
                return
        except UnicodeDecodeError:
            pass

        # If UTF-8 fails, try reading as binary and handle different encodings
        with open(log_file_path, 'rb') as log_file:
            content = log_file.read()
            
            # Try different encodings based on BOM
            if content.startswith(b'\xff\xfe'):  # UTF-16 LE BOM
                decoded = content[2:].decode('utf-16-le')
            elif content.startswith(b'\xfe\xff'):  # UTF-16 BE BOM
                decoded = content[2:].decode('utf-16-be')
            elif content.startswith(b'\xef\xbb\xbf'):  # UTF-8 BOM
                decoded = content[3:].decode('utf-8')
            else:
                # Try common encodings
                for encoding in ['cp1252', 'latin1', 'iso-8859-1', 'utf-16', 'ascii']:
                    try:
                        decoded = content.decode(encoding)
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    # If all fail, use utf-8 with error handling
                    decoded = content.decode('utf-8', errors='replace')
            
            # Clean up any BOM or invalid characters at the start
            decoded = decoded.lstrip('\ufeff\ufffe')
            print(decoded)

    except FileNotFoundError:
        print(f"Error: The file '{log_file_path}' does not exist.")
    except Exception as e:
        print(f"Error reading log file: {str(e)}")

if __name__ == "__main__":
    # Check if a log file path was provided as an argument
    if len(sys.argv) != 2:
        print("Usage: python tests/print_specific_log.py <log_file_path>")
        sys.exit(1)

    log_file_path = sys.argv[1]

    # If it's an absolute path, use it directly
    if os.path.isabs(log_file_path):
        print_log_content(log_file_path)
    else:
        # If it's a relative path, try multiple possible locations
        possible_paths = [
            log_file_path,  # Try as-is first
            os.path.join(os.getcwd(), log_file_path),  # Try from current directory
            os.path.join(os.getcwd(), 'tests', log_file_path),  # Try in tests folder
        ]

        # Try each possible path
        for path in possible_paths:
            if os.path.exists(path):
                print_log_content(path)
                break
        else:
            print(f"Error: Could not find log file in any of these locations:")
            for path in possible_paths:
                print(f"- {path}")