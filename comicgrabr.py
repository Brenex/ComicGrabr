#!/usr/bin/env python3
"""
Comic Grabber Bot

This script automates the process of managing a comic pull list and downloading released comics.
It performs the following key functions:
1.  **League of Comic Geeks (LCG) Integration**: Logs into your LCG account to download your latest pull list in Excel format.
2.  **Local Pull List Management**: Parses the downloaded Excel file and updates a local JSON database ('pull_list.json') with upcoming comic releases, while removing past releases.
3.  **AirDC++ Search and Download**: For comics released on the current day (or specified past dates via argument), it searches for corresponding files on AirDC++ hubs and queues them for download.
4.  **Discord Notifications**: Sends detailed notifications to a configured Discord webhook about script status (start, completion, errors), queued downloads, and upcoming comic releases for the next Wednesday.
5.  **Logging**: Implements a robust logging system that outputs high-level information to the console and detailed debug information to timestamped log files, with automatic cleanup of old logs.
6.  **Command-Line Arguments**: Supports various operational modes, including processing a specific Excel file (for initial setup or manual updates), searching for past releases, and performing dry runs without initiating actual downloads.

This bot is designed to streamline your comic collection process, ensuring you stay up-to-date with your favorite series.
"""
# Standard library imports
import argparse  # For parsing command-line arguments
import base64  # For encoding/decoding in Base64 (e.g., for AirDC++ auth)
import json  # For file operations using JSON
import logging  # For logging functionality
import os  # For interacting with the operating system
import sys  # For interacting with the system (e.g., stdout for logging)
import time  # For delays and timing
from datetime import datetime, timedelta  # For handling dates and times

# Third-party imports
import pandas as pd  # For data manipulation and analysis
import requests  # For making HTTP requests
import xlrd  # For reading Excel files (advanced options)
from bs4 import BeautifulSoup  # For parsing HTML/XML
from discord_webhook import (
    DiscordWebhook,
    DiscordEmbed,
)  # For sending Discord notifications
from dotenv import load_dotenv  # For loading environment variables from a .env file

# --- CONFIGURATION START ---
# Load environment variables from .env file
load_dotenv()

# Define log retention for the script's own log files
LOG_RETENTION_DAYS = 7

# Default LOG_LEVEL. This will be overridden by command-line argument if provided.
DEFAULT_LOG_LEVEL = logging.INFO

# Default qBittorrent connection details (can be overridden by command-line arguments or environment variables)
# Prioritize environment variables, then command-line defaults
# (These variables are not used by the comic grabber logic, but kept for consistency if part of a larger automation suite)
DEFAULT_QB_HOST = os.getenv("QB_HOST", "http://localhost:8080")
DEFAULT_QB_USER = os.getenv("QB_USER", "admin")
DEFAULT_QB_PASSWORD = os.getenv("QB_PASSWORD", "adminadmin")

# Discord Webhook URL for notifications (optional, prioritized from environment)
DISCORD_WEBHOOK_URL = os.getenv("DISCORD_WEBHOOK_URL", "")

# --- Comic Grabber Specific Configuration ---
LCG_USERNAME = os.getenv("LCG_USERNAME")
LCG_PASSWORD = os.getenv("LCG_PASSWORD")
AIRDCPP_API_URL = os.getenv(
    "AIRDCPP_API_URL"
)  # Should be "http://127.0.0.1:5600/api/v1/"
AIRDCPP_USERNAME = os.getenv("AIRDCPP_USERNAME")
AIRDCPP_PASSWORD = os.getenv("AIRDCPP_PASSWORD")
AIRDCPP_INSTANCE_ID = os.getenv(
    "AIRDCPP_INSTANCE_ID"
)  # Note: This is not currently used, dynamic instance IDs are preferred.

LOGIN_URL = "https://leagueofcomicgeeks.com/login"
EXPORT_URL = "https://leagueofcomicgeeks.com/member/export_pulls"
OUTPUT_FILENAME = "league_of_comic_geeks_pulls.xls"  # Temporary LCG export file
PULL_LIST_DB_FILE = "pull_list.json"  # The JSON file to store your pull list
# DOWNLOADED_TTHS_FILE = "downloaded_tths.json" # Removed - No longer needed

# User-Agent header to mimic a web browser
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

# Global variable to store the Bearer token
AIRDCPP_AUTH_TOKEN = None

# Global set to store TTHs of downloaded comics for quick lookup
# DOWNLOADED_TTHS = set() # Removed - No longer needed

# --- CONFIGURATION END ---

# --- LOGGING CONFIGURATION START ---
# Create a logger instance
logger = logging.getLogger(__name__)
# Log level will be set dynamically based on CLI args in main()
# logger.setLevel(LOG_LEVEL) # Removed from here

# Create a formatter for log messages
formatter = logging.Formatter(
    "%(asctime)s - %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
)

# Console handler (outputs to stdout)
console_handler = logging.StreamHandler(sys.stdout)
# console_handler.setLevel(logging.INFO) # Level set dynamically in main()
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)


def get_logs_dir():
    """
    Returns the path to the logs directory, creating it if it doesn't exist.

    Returns:
        str: The absolute path to the logs directory.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    logs_dir = os.path.join(script_dir, "logs")
    os.makedirs(logs_dir, exist_ok=True)
    return logs_dir


def get_current_run_log_file_path():
    """
    Generates a timestamped log file path for the current script run.

    Returns:
        str: The absolute path to the log file for the current run.
    """
    logs_dir = get_logs_dir()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return os.path.join(
        logs_dir, f"comic_grabber_bot_{timestamp}.log"
    )  # Unified log file name for this script


def cleanup_old_logs():
    """
    Deletes log files in the 'logs' directory that are older than LOG_RETENTION_DAYS.
    Targets files named 'comic_grabber_bot_*.log'.
    """
    logs_dir = get_logs_dir()
    cutoff_time = datetime.now() - timedelta(days=LOG_RETENTION_DAYS)

    logger.info(
        f"Cleaning up old logs in '{logs_dir}' older than {LOG_RETENTION_DAYS} days."
    )

    for filename in os.listdir(logs_dir):
        # Target files starting with 'comic_grabber_bot_' and ending with '.log'
        if filename.startswith("comic_grabber_bot_") and filename.endswith(".log"):
            file_path = os.path.join(logs_dir, filename)
            try:
                file_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                if file_time < cutoff_time:
                    os.remove(file_path)
                    logger.info(f"Removed old log file: {filename}")
            except OSError as e:
                logger.error(
                    f"Failed to delete old log file {filename}: {e}", exc_info=True
                )
            except Exception as e:
                logger.error(
                    f"An unexpected error occurred while processing log file {filename}: {e}",
                    exc_info=True,
                )


# Get the log file path for the current run
current_script_log_file_path = get_current_run_log_file_path()

# File handler (outputs to a file)
file_handler = None # Initialize to None
try:
    file_handler = logging.FileHandler(current_script_log_file_path)
    # file_handler.setLevel(LOG_LEVEL) # Level set dynamically in main()
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    logger.info(f"Logging current run to: {current_script_log_file_path}")
except Exception as e:
    logger.error(f"Failed to set up file logger at {current_script_log_file_path}: {e}")
    logger.warning("Continuing with console-only logging due to file logging error.")

# --- LOGGING CONFIGURATION END ---


# --- Helper Functions (Discord Notifications) ---


def send_discord_notification(
    webhook_url: str,
    title: str,
    description: str,
    color: int,
    fields: list = None,
    log_file_path: str = None,
    is_dry_run: bool = False,
):
    """
    Sends a rich Discord embed notification, optionally with a log file attachment.

    Args:
        webhook_url (str): The Discord webhook URL.
        title (str): The title of the embed.
        description (str): The main description of the embed.
        color (int): The color of the embed sidebar (e.g., 0x00FF00 for green).
        fields (list, optional): A list of field dictionaries for the embed. Defaults to None.
        log_file_path (str, optional): Path to a log file to attach. Defaults to None.
        is_dry_run (bool, optional): If True, prefixes title and description as a dry run. Defaults to False.
    """
    if not webhook_url:
        logger.warning(
            "Discord webhook URL is not configured. Skipping Discord notification."
        )
        return

    # Apply dry run prefixes
    if is_dry_run:
        title = f"[DRY RUN] {title}"
        description = f"**This is a dry run. No actual downloads were initiated or state changed.**\n\n{description}"
        color = 0xAAAAAA  # Grey color for dry runs

    payload = {
        "embeds": [
            {
                "title": title,
                "description": description,
                "color": color,
                "timestamp": datetime.now().isoformat(),
                "fields": fields if fields else [],
                "footer": {"text": "Comic Grabber Bot"},
            }
        ]
    }

    files = {}
    if log_file_path and os.path.exists(log_file_path):
        try:
            files = {
                "file": (
                    os.path.basename(log_file_path),
                    open(log_file_path, "rb"),
                    "text/plain",
                )
            }
        except Exception as e:
            logger.error(f"Failed to open log file for Discord upload: {e}")
            log_file_path = None  # Don't try to send file if opening failed

    logger.debug(f"Attempting to send Discord notification to: {webhook_url}")
    try:
        if files:
            # When sending files, the payload must be sent as a separate 'payload_json' part
            # and the 'Content-Type' is handled by requests with 'multipart/form-data'
            response = requests.post(
                webhook_url,
                data={"payload_json": requests.utils.quote(json.dumps(payload))},
                files=files,
            )
        else:
            response = requests.post(webhook_url, json=payload)

        response.raise_for_status()  # Raise an exception for HTTP errors (4xx or 5xx)
        logger.info("Successfully sent Discord notification.")
    except requests.exceptions.HTTPError as errh:
        logger.error(f"Discord HTTP Error: {errh} - {errh.response.text}")
    except requests.exceptions.ConnectionError as errc:
        logger.error(f"Discord Connection Error: {errc}")
    except requests.exceptions.Timeout as errt:
        logger.error(f"Discord Timeout Error: {errt}")
    except requests.exceptions.RequestException as err:
        logger.error(f"Discord Request Error: {err}")
    except Exception as e:
        logger.error(
            f"An unexpected error occurred while sending Discord notification: {e}",
            exc_info=True,
        )
    finally:
        if (
            "file" in files and files["file"] and files["file"][1]
        ):  # Check if file was opened before trying to close
            files["file"][1].close()  # Ensure file handle is closed


# --- League of Comic Geeks Interaction ---


def login_and_download_pull_list():
    """
    Logs into League of Comic Geeks, downloads the pull list Excel file,
    and returns its path.

    Returns:
        str or None: The path to the downloaded Excel file if successful, otherwise None.
    """
    if not LCG_USERNAME or not LCG_PASSWORD:
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: LCG Credentials Missing",
            description="LCG_USERNAME or LCG_PASSWORD not set. Cannot download pull list.",
            color=0xFF0000,
        )
        return None

    session = requests.Session()
    session.headers.update(HEADERS)

    try:
        logger.debug(f"Fetching login page from {LOGIN_URL}...")
        login_page_response = session.get(LOGIN_URL)
        login_page_response.raise_for_status()

        soup = BeautifulSoup(login_page_response.text, "html.parser")
        csrf_token_tag = soup.find("input", {"name": "ci_csrf_token"})

        if csrf_token_tag and "value" in csrf_token_tag.attrs:
            csrf_token = csrf_token_tag["value"]
            logger.debug(f"Extracted CSRF token.")
        else:
            logger.error("CSRF token not found. Website structure might have changed.")
            send_discord_notification(
                webhook_url=DISCORD_WEBHOOK_URL,
                title="Error: LCG CSRF Token",
                description="LCG CSRF token not found during login. Website structure might have changed.",
                color=0xFF0000,
            )
            return None

        payload = {
            "username": LCG_USERNAME,
            "password": LCG_PASSWORD,
            "ci_csrf_token": csrf_token,
            "submit": "Continue »",
        }

        logger.debug(f"Attempting to log in to {LOGIN_URL}...")
        login_response = session.post(LOGIN_URL, data=payload, allow_redirects=True)
        login_response.raise_for_status()

        if "My Comics" in login_response.text or "member" in login_response.url:
            logger.info("Successfully logged in to League of Comic Geeks!")
            time.sleep(2)

            logger.debug(f"Attempting to download Excel file from {EXPORT_URL}...")
            excel_response = session.get(EXPORT_URL, stream=True)
            excel_response.raise_for_status()

            if excel_response.status_code == 200:
                os.makedirs(os.path.dirname(OUTPUT_FILENAME) or ".", exist_ok=True)
                with open(OUTPUT_FILENAME, "wb") as f:
                    for chunk in excel_response.iter_content(chunk_size=8192):
                        f.write(chunk)
                logger.info(f"Excel file downloaded successfully to {OUTPUT_FILENAME}")
                return OUTPUT_FILENAME
            else:
                send_discord_notification(
                    webhook_url=DISCORD_WEBHOOK_URL,
                    title="Error: Excel Download Failed",
                    description=f"Failed to download Excel file from LCG. Status code: {excel_response.status_code}",
                    color=0xFF0000,
                )
                return None
        else:
            logger.error("Login failed. Check LCG username/password.")
            send_discord_notification(
                webhook_url=DISCORD_WEBHOOK_URL,
                title="Error: LCG Login Failed",
                description="LCG login failed. Check username/password.",
                color=0xFF0000,
            )
            return None

    except requests.exceptions.RequestException as e:
        logger.error(f"A network or HTTP error occurred during LCG interaction: {e}")
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error during LCG Interaction",
            description=f"A network or HTTP error occurred during LCG login/download: {e}",
            color=0xFF0000,
        )
        return None
    except Exception as e:
        logger.error(
            f"An unexpected error occurred during LCG interaction: {e}", exc_info=True
        )
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Unexpected Error during LCG Interaction",
            description=f"An unexpected error occurred during LCG interaction: {e}",
            color=0xFF0000,
        )
        return None


def update_json_pull_list_from_excel(excel_file_path):
    """
    Parses the Excel file (Pulls sheet, Comic and Release columns)
    and updates the JSON pull list file with future releases.

    Args:
        excel_file_path (str): The path to the downloaded Excel pull list file.

    Returns:
        bool: True if the JSON pull list was successfully updated, False otherwise.
    """
    if not os.path.exists(excel_file_path):
        logger.error(f"Error: Excel file not found at {excel_file_path}")
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: Excel File Not Found",
            description=f"Excel pull list file not found at {excel_file_path}. Cannot update JSON.",
            color=0xFF0000,
        )
        return False

    logger.info(
        f"Processing Excel file '{excel_file_path}' to update JSON pull list..."
    )
    df = None
    try:
        logger.debug(
            "Attempting to parse Excel file with xlrd (for .xls format), ignoring corruption..."
        )
        workbook = xlrd.open_workbook_xls(
            excel_file_path, ignore_workbook_corruption=True
        )
        df = pd.read_excel(workbook, sheet_name="Pulls", engine="xlrd")
    except Exception as e:
        logger.critical(f"Error reading Excel file or 'Pulls' sheet: {e}.")
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Critical Error: Excel Read Failure",
            description=f"Could not read Excel file or 'Pulls' sheet from '{excel_file_path}'. Details: {e}",
            color=0xFF0000,
        )
        return False

    if df is None:
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: Excel Data Load Failed",
            description=f"Failed to load Excel data from '{excel_file_path}'.",
            color=0xFF0000,
        )
        return False

    detected_columns = df.columns.tolist()
    logger.debug(f"Excel file columns detected: {detected_columns}")

    if "Comic" not in detected_columns or "Release" not in detected_columns:
        logger.error(
            "Error: Missing 'Comic' or 'Release' column in Excel 'Pulls' sheet."
        )
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: Missing Excel Columns",
            description="Excel 'Pulls' sheet missing 'Comic' or 'Release' column. Cannot update JSON.",
            color=0xFF0000,
        )
        return False

    # Initialize a new map to store comics. This effectively "clears" previous data
    # as we will only populate this with data from the current Excel file.
    new_comics_map = {}
    comics_added_or_updated = 0
    today = datetime.now().date()

    for index, row in df.iterrows():
        try:
            comic_name_raw = row.get("Comic")
            release_date_raw = row.get("Release")

            if pd.isna(comic_name_raw) or pd.isna(release_date_raw):
                logger.debug(
                    f"Skipping row {index}: Missing comic name or release date."
                )
                continue

            comic_name = (
                str(comic_name_raw).replace("#", "").replace(":", "").strip()
            )  # Remove '#' and ':' from comic name
            release_date = None

            if isinstance(release_date_raw, datetime):
                release_date = release_date_raw.date()
            else:
                try:
                    release_date = datetime.strptime(
                        str(release_date_raw), "%Y-%m-%d"
                    ).date()
                    logger.debug(
                        f"Parsed date '{release_date_raw}' as %Y-%m-%d for '{comic_name}'."
                    )
                except ValueError:
                    try:
                        release_date = datetime.strptime(
                            str(release_date_raw), "%m/%d/%Y"
                        ).date()
                        logger.debug(
                            f"Parsed date '{release_date_raw}' as %m/%d/%Y for '{comic_name}'."
                        )
                    except ValueError:
                        logger.warning(
                            f"Warning: Could not parse release date '{release_date_raw}' for comic '{comic_name}'. Skipping."
                        )
                        continue  # Skip this comic if date can't be parsed

            comic_data = {
                "comic_name": comic_name,
                "release_date": release_date.strftime("%Y-%m-%d"),  # Store as string
            }
            key = f"{comic_name}-{comic_data['release_date']}"

            # Only store future or today's releases in the new JSON file
            if release_date >= today:
                if key not in new_comics_map:
                    comics_added_or_updated += 1
                    logger.debug(
                        f"Adding new comic to JSON: {comic_name} ({comic_data['release_date']})"
                    )
                else:
                    logger.debug(
                        f"Updating existing comic in JSON: {comic_name} ({comic_data['release_date']})"
                    )
                new_comics_map[key] = comic_data  # Add or update
            # No need to explicitly remove past comics, as new_comics_map starts empty
            # and only future/today's comics are added.

        except Exception as e:
            logger.warning(
                f"Warning: Could not process row {index} for JSON update. Error: {e}. Row data: {row.to_dict()}"
            )

    # Convert new map back to list
    updated_comics_list = list(new_comics_map.values())

    # Sort the list by 'release_date'
    updated_comics_list.sort(key=lambda comic: comic["release_date"])

    try:
        # Writing mode 'w' truncates the file if it exists or creates a new one.
        with open(PULL_LIST_DB_FILE, "w", encoding="utf-8") as f:
            json.dump(updated_comics_list, f, indent=4, ensure_ascii=False)
        logger.info(
            f"Successfully updated/added {comics_added_or_updated} comics to {PULL_LIST_DB_FILE}."
        )
        return True
    except IOError as e:
        logger.error(f"Error writing to {PULL_LIST_DB_FILE}: {e}")
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: JSON Write Failed",
            description=f"Error writing to JSON pull list file: {e}",
            color=0xFF0000,
        )
        return False


# --- AirDC++ Interaction ---


def get_bearer_token(is_dry_run=False):
    """
    Obtains a Bearer token from the AirDC++ API for authentication.

    Args:
        is_dry_run (bool, optional): If True, indicates a dry run, affecting notifications. Defaults to False.

    Returns:
        str or None: The Bearer token string if successful, otherwise None.
    """
    global AIRDCPP_AUTH_TOKEN
    if AIRDCPP_AUTH_TOKEN:  # Reuse token if already obtained
        logger.debug("Reusing existing AirDC++ Bearer token.")
        return AIRDCPP_AUTH_TOKEN

    if not AIRDCPP_API_URL or not AIRDCPP_USERNAME or not AIRDCPP_PASSWORD:
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: AirDC++ Credentials Missing",
            description="AirDC++ API URL or credentials not set. Cannot obtain Bearer token.",
            color=0xFF0000,
            is_dry_run=is_dry_run,
        )
        return None

    auth_endpoint = f"{AIRDCPP_API_URL}sessions/authorize"
    auth_data = {
        "username": AIRDCPP_USERNAME,
        "password": AIRDCPP_PASSWORD,
        "max_inactivity": 3600,  # Token valid for 1 hour of inactivity
    }

    logger.debug(f"Attempting to obtain Bearer token from {auth_endpoint}...")
    try:
        response = requests.post(auth_endpoint, json=auth_data, timeout=10)
        response.raise_for_status()
        auth_response = response.json()

        if "auth_token" in auth_response:
            AIRDCPP_AUTH_TOKEN = auth_response["auth_token"]
            logger.info("Successfully obtained AirDC++ Bearer token.")
            return AIRDCPP_AUTH_TOKEN
        else:
            logger.error("Bearer token not found in AirDC++ authorization response.")
            send_discord_notification(
                webhook_url=DISCORD_WEBHOOK_URL,
                title="Error: Bearer Token Missing",
                description="Bearer token not found in AirDC++ authorization response.",
                color=0xFF0000,
                is_dry_run=is_dry_run,
            )
            return None
    except requests.exceptions.Timeout:
        logger.warning(f"AirDC++ token authorization timed out.")
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Warning: AirDC++ Timeout",
            description=f"AirDC++ token authorization timed out.",
            color=0xFF8C00,
            is_dry_run=is_dry_run,
        )
        return None
    except requests.exceptions.RequestException as e:
        logger.error(f"Error obtaining AirDC++ Bearer token: {e}")
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: AirDC++ Token Failed",
            description=f"Error obtaining AirDC++ Bearer token: {e}",
            color=0xFF0000,
            is_dry_run=is_dry_run,
        )
        return None


def get_airdcpp_auth_headers(is_dry_run=False):
    """
    Returns authentication headers for AirDC++ API using the Bearer token.

    Args:
        is_dry_run (bool, optional): If True, indicates a dry run, affecting token acquisition notifications. Defaults to False.

    Returns:
        dict: A dictionary containing the Authorization header, or an empty dictionary if authentication fails.
    """
    token = get_bearer_token(is_dry_run)
    if token:
        return {"Authorization": f"Bearer {token}"}
    return {}


def search_airdcpp(comic_name, is_dry_run=False):
    """
    Performs the three-step AirDC++ search process:
    1. Creates a search instance.
    2. Executes the hub search.
    3. Retrieves results with retries.

    Args:
        comic_name (str): The name of the comic to search for.
        is_dry_run (bool, optional): If True, indicates a dry run, affecting notifications. Defaults to False.

    Returns:
        tuple: A tuple containing (dict, str) if a suitable match is found,
               where the dict includes 'id', 'name', 'path', 'size', 'tth' of the best match,
               and the str is the session_search_id.
               Returns (None, str) if no match is found but session_search_id was obtained.
               Returns (None, None) if search instance creation failed.
    """
    if not AIRDCPP_API_URL:
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: AirDC++ URL Missing",
            description="AIRDCPP_API_URL not set. Cannot perform AirDC++ search.",
            color=0xFF0000,
            is_dry_run=is_dry_run,
        )
        return None, None  # Return None for both match and session_search_id

    headers = get_airdcpp_auth_headers(is_dry_run)
    if not headers:
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: AirDC++ Auth Failed",
            description="AirDC++ authentication failed. Cannot search.",
            color=0xFF0000,
            is_dry_run=is_dry_run,
        )
        return None, None  # Return None for both match and session_search_id

    session_search_id = None
    hub_search_operation_id = None

    # --- Step 1: Create a search instance (POST /api/v1/search) ---
    search_instance_create_endpoint = f"{AIRDCPP_API_URL}search"
    initial_instance_payload = {
        "pattern": comic_name,
        "limit": 10,
        "expiration": 5,
    }  # 5 minutes expiration

    logger.debug(f"  Attempting to create search instance for '{comic_name}'...")
    try:
        response = requests.post(
            search_instance_create_endpoint,
            json=initial_instance_payload,
            headers=headers,
            timeout=15,
        )
        response.raise_for_status()
        response_json = response.json()
        if "id" in response_json:
            session_search_id = response_json["id"]
            logger.info(
                f"  Successfully created search instance. Session Search ID: {session_search_id}"
            )
        else:
            logger.error(
                "  No 'id' found in search instance creation response. Cannot proceed."
            )
            send_discord_notification(
                webhook_url=DISCORD_WEBHOOK_URL,
                title="Error: AirDC++ Search Instance Failed",
                description=f"Error creating AirDC++ search instance for '{comic_name}'. No ID found in response.",
                color=0xFF8C00,
                is_dry_run=is_dry_run,
            )
            return None, None
    except requests.exceptions.RequestException as e:
        logger.error(f"  Error creating search instance for '{comic_name}': {e}")
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: AirDC++ Search Instance Creation",
            description=f"Error creating AirDC++ search instance for '{comic_name}': {e}",
            color=0xFF0000,
            is_dry_run=is_dry_run,
        )
        return None, None

    # --- Step 2: Perform the Hub Search (POST /api/v1/search/{instance_id}/hub_search) ---
    hub_search_execute_endpoint = (
        f"{AIRDCPP_API_URL}search/{session_search_id}/hub_search"
    )

    # The 'query' field itself is an object, containing 'pattern' and optional 'file_extensions'.
    hub_search_payload_base = {"pattern": comic_name, "limit": 10}

    hub_search_payloads_to_try = [
        {"query": {**hub_search_payload_base, "file_extensions": ["cbz"]}},
        {"query": {**hub_search_payload_base, "file_extensions": ["cbr"]}},
        {"query": hub_search_payload_base},  # General search fallback
    ]

    logger.debug(
        f"  Performing hub search for '{comic_name}' (using Session ID: {session_search_id})..."
    )
    for payload_attempt in hub_search_payloads_to_try:
        try:
            time.sleep(
                2
            )  # Give a moment between creating instance and performing search
            response = requests.post(
                hub_search_execute_endpoint,
                json=payload_attempt,
                headers=headers,
                timeout=20,
            )
            response.raise_for_status()
            response_json = response.json()

            if "search_id" in response_json:
                hub_search_operation_id = response_json["search_id"]
                logger.debug(
                    f"  Hub search initiated successfully. Hub Search Operation ID: {hub_search_operation_id}"
                )
                break  # Success, proceed to step 3
            else:
                logger.debug(
                    f"  No 'search_id' found in hub search response for payload: {json.dumps(payload_attempt)}. Trying next."
                )

        except requests.exceptions.RequestException as e:
            logger.error(
                f"  Error during hub search execution with payload: {json.dumps(payload_attempt)}: {e}"
            )
            send_discord_notification(
                webhook_url=DISCORD_WEBHOOK_URL,
                title="Error: AirDC++ Hub Search Execution",
                description=f"Error during AirDC++ hub search for '{comic_name}' with payload {json.dumps(payload_attempt)}: {e}",
                color=0xFF0000,
                is_dry_run=is_dry_run,
            )
            continue

    if not hub_search_operation_id:
        logger.error(
            "  Failed to initiate hub search. No hub search operation ID obtained. Cannot proceed to get results."
        )
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: AirDC++ Hub Search Failed",
            description=f"Failed to initiate AirDC++ hub search for '{comic_name}'. No operation ID obtained.",
            color=0xFF8C00,
            is_dry_run=is_dry_run,
        )
        return (
            None,
            session_search_id,
        )  # Return session_search_id even if hub search failed

    # --- Step 3: Retrieve Results (GET /api/v1/search/{session_id}/results/start/count) with Retries ---
    start_index = 0
    count_limit = 100  # Fetch up to 100 results
    max_retries = 3
    initial_delay = 7  # seconds
    delay_increment = 5  # seconds

    # Endpoint expects the session ID from the *first* step
    results_fetch_endpoint = f"{AIRDCPP_API_URL}search/{session_search_id}/results/{start_index}/{count_limit}"

    logger.debug(
        f"  Fetching AirDC++ search results for '{comic_name}' (using Session Search ID: {session_search_id})..."
    )

    results = []  # Initialize results as an empty list

    for attempt in range(max_retries):
        current_delay = initial_delay + (attempt * delay_increment)
        logger.debug(
            f"  Waiting {current_delay} seconds before fetching results (Attempt {attempt + 1}/{max_retries})..."
        )
        time.sleep(current_delay)

        try:
            response = requests.get(results_fetch_endpoint, headers=headers, timeout=30)
            response.raise_for_status()
            results = response.json()

            # If results are found (list not empty), break the retry loop
            if isinstance(results, list) and results:
                logger.debug(f"  Results found on attempt {attempt + 1}.")
                break
            else:
                logger.debug(
                    f"  No results found on attempt {attempt + 1}. Retrying..."
                )
                results = []  # Ensure it's an empty list for clarity and loop condition
        except requests.exceptions.Timeout:
            logger.warning(
                f"  AirDC++ results fetching for '{comic_name}' timed out on attempt {attempt + 1}."
            )
            results = []  # Reset results for retry
        except requests.exceptions.RequestException as e:
            logger.error(
                f"  Error fetching AirDC++ search results for '{comic_name}' on attempt {attempt + 1}: {e}"
            )
            send_discord_notification(
                webhook_url=DISCORD_WEBHOOK_URL,
                title="Error: AirDC++ Results Fetching",
                description=f"Error fetching AirDC++ search results for '{comic_name}' on attempt {attempt + 1}: {e}",
                color=0xFF0000,
                is_dry_run=is_dry_run,
            )
            results = []  # Reset results for retry
            # If it's a non-retriable error (e.g., 404), maybe break here? For now, keep retrying.

    if isinstance(results, list) and results:
        comic_files = [
            r
            for r in results
            if r.get("path")
            and (
                r["path"].lower().endswith(".cbz") or r["path"].lower().endswith(".cbr")
            )
        ]

        if comic_files:
            # Prioritize .cbz over .cbr if both exist, otherwise pick first available
            best_match = next(
                (f for f in comic_files if f["path"].lower().endswith(".cbz")), None
            )
            if not best_match:
                best_match = comic_files[0]  # Fallback to first if no cbz
            
            logger.info(
                f"  Found match: {best_match.get('path')} (ID: {best_match.get('id')})"
            )
            return {
                "id": best_match.get("id"),
                "name": best_match.get("name"),  # Include name for target_name
                "path": best_match.get("path"),
                "size": best_match.get("size"),
                "tth": best_match.get("tth"),  # Ensure tth is included for download
            }, session_search_id  # Return the match and the session_search_id
        else:
            logger.info(
                f"  No .cbz/.cbr file found among search results for '{comic_name}'."
            )
    else:
        logger.info(f"  No search results found on AirDC++ for '{comic_name}'.")

    return None, session_search_id  # Always return session_search_id even if no match


def download_airdcpp(file_info, session_search_id, is_dry_run=False):
    """
    Initiates a download on AirDC++ given file_info (containing id, name, size, tth)
    and the session_search_id used for the original search to retrieve file details.

    Args:
        file_info (dict): A dictionary containing 'id', 'name', 'size', and 'tth' of the file to download.
        session_search_id (str): The session ID from the initial search instance creation, required by AirDC++ for some detailed lookups.
        is_dry_run (bool, optional): If True, indicates a dry run, affecting notifications. Defaults to False.

    Returns:
        bool: True if the download command was successfully sent, "skipped" if already exists, False otherwise.
    """
    if is_dry_run:
        logger.info(
            f"  [DRY RUN] Would queue download for file ID: {file_info['id']} (Name: {file_info['name']})."
        )
        return True  # Simulate success for dry run

    headers = get_airdcpp_auth_headers(is_dry_run)
    if not headers:
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: AirDC++ Auth Failed",
            description="AirDC++ authentication failed. Cannot download.",
            color=0xFF0000,
            is_dry_run=is_dry_run,
        )
        return False

    # Extract required fields for download bundle directly from file_info
    target_name = file_info.get("name")
    size = file_info.get("size")
    tth = file_info.get("tth")

    if target_name is None or size is None or tth is None:
        logger.error(
            f"  Error: Missing 'name', 'size', or 'tth' in file info for {file_info['id']}."
        )
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: Missing File Details for Download",
            description=f"Missing required file details (name, size, or TTH) for {file_info.get('path', 'unknown comic')}.",
            color=0xFF0000,
            is_dry_run=is_dry_run,
        )
        return False

    # --- Queue download bundle (POST /api/v1/queue/bundles/file) ---
    download_bundle_endpoint = f"{AIRDCPP_API_URL}queue/bundles/file"
    download_data = {"target_name": target_name, "size": size, "tth": tth}

    logger.debug(
        f"Attempting to queue download bundle to {download_bundle_endpoint} for '{target_name}'..."
    )
    try:
        response = requests.post(
            download_bundle_endpoint, json=download_data, headers=headers, timeout=10
        )
        response.raise_for_status()
        logger.info(f"  Successfully sent download command for '{target_name}'.")
        return True

    except requests.exceptions.Timeout:
        logger.warning(f"  AirDC++ download queue for '{target_name}' timed out.")
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Warning: AirDC++ Download Timeout",
            description=f"AirDC++ download queue for '{target_name}' timed out.",
            color=0xFF8C00,
            is_dry_run=is_dry_run,
        )
        return False
    except requests.exceptions.RequestException as e:
        # Check for the specific "File exists" message in the response text
        if e.response is not None and "File exists on the disk already" in e.response.text:
            logger.info(
                f"  Download skipped for '{target_name}': File already exists on disk or in queue. (TTH: {tth})"
            )
            send_discord_notification(
                webhook_url=DISCORD_WEBHOOK_URL,
                title="Comic Already Exists",
                description=f"**Skipped:** {target_name}\nThis comic is already on disk or in AirDC++ queue. (TTH: `{tth}`)",
                color=0xFFFF00,  # Yellow for warning/info
                is_dry_run=is_dry_run,
            )
            return "skipped"  # Return "skipped" to indicate this specific scenario
        else:
            logger.error(f"  Error initiating AirDC++ download for '{target_name}': {e}")
            send_discord_notification(
                webhook_url=DISCORD_WEBHOOK_URL,
                title="Error: AirDC++ Download Failed",
                description=f"Error queuing AirDC++ download for '{target_name}': {e}",
                color=0xFF0000,
                is_dry_run=is_dry_run,
            )
            return False


# --- Main Automation Logic ---

def log_level_type(arg):
    """Custom type function for argparse to convert log level input to uppercase."""
    return arg.upper()

def main():
    """
    Main function to run the Comic Grabber Bot.
    It handles command-line arguments, logs in to League of Comic Geeks,
    updates the local JSON pull list, searches and downloads comics from AirDC++,
    and sends Discord notifications for various events.
    """
    script_start_time = datetime.now()

    # Setup command-line argument parsing
    parser = argparse.ArgumentParser(
        description="Monitors League of Comic Geeks pull list, stores future releases in a JSON file, and downloads today's releases from AirDC++."
    )
    parser.add_argument(
        "--excel-file",
        "-f",
        type=str,
        help="Path to a previously downloaded LCG pull list Excel file. If provided, "
        "the script will read this file to update the JSON database and then exit. "
        "It will NOT perform daily AirDC++ searches in this mode.",
    )
    parser.add_argument(
        "--search-past-releases",
        action="store_true",
        help="Search for all comics currently listed in the JSON file with a release date today or in the past.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Perform a dry run: search for comics but do not initiate actual downloads or send live Discord notifications (notifications will be marked as dry run).",
    )
    parser.add_argument(
        "--log-level",
        type=log_level_type, # Use the custom type function here
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
        help="Set the logging level (e.g., INFO, DEBUG, WARNING). Default is INFO.",
    )
    args = parser.parse_args()

    # Set logging level based on command-line argument
    numeric_log_level = getattr(logging, args.log_level, DEFAULT_LOG_LEVEL) # args.log_level is already uppercase now
    logger.setLevel(numeric_log_level)
    console_handler.setLevel(numeric_log_level)
    if file_handler: # Only set if file_handler was successfully created
        file_handler.setLevel(numeric_log_level)


    # Clean up old logs at the start of the main execution
    cleanup_old_logs()

    # Get current day of the week (Monday is 0, Wednesday is 2)
    today_weekday = datetime.now().weekday()
    is_wednesday = today_weekday == 2  # 2 represents Wednesday

    # Send script started notification
    logger.info("--- Starting Comic Grabber Bot ---")
    send_discord_notification(
        webhook_url=DISCORD_WEBHOOK_URL,
        title="Comic Grabber Bot",
        description=f"Script execution commenced at: {script_start_time.strftime('%Y-%m-%d %H:%M:%S')}",
        color=0x3498DB,  # Blue color for informational start
        is_dry_run=args.dry_run,
    )

    if not is_wednesday and not args.excel_file: # Only skip download logic if NOT Wednesday AND NOT explicitly running with --excel-file
        logger.info(
            f"Today is not Wednesday ({datetime.now().strftime('%A')}). Downloading and updating pull list only."
        )
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Pull List Sync Only",
            description=f"Today is {datetime.now().strftime('%A')}. Only updating the comic pull list. No downloads will be attempted.",
            color=0x00BFFF,  # Deep Sky Blue
            is_dry_run=args.dry_run,
        )
        
        pulled_comics_source_file = login_and_download_pull_list()
        if pulled_comics_source_file:
            json_update_success = update_json_pull_list_from_excel(
                pulled_comics_source_file
            )
            if json_update_success:
                logger.info("Pull list downloaded and JSON file updated successfully.")
                
                # Load comics from the JSON file for next Wednesday's release check
                all_comics_from_json_for_upcoming = [] 
                if os.path.exists(PULL_LIST_DB_FILE):
                    try:
                        with open(PULL_LIST_DB_FILE, "r", encoding="utf-8") as f:
                            all_comics_from_json_for_upcoming = json.load(f)
                    except (json.JSONDecodeError, FileNotFoundError) as e:
                        logger.error(f"Error reading or parsing {PULL_LIST_DB_FILE} for upcoming releases: {e}.")
                        send_discord_notification(
                            webhook_url=DISCORD_WEBHOOK_URL,
                            title="Error: JSON Read Failed for Upcoming",
                            description=f"Error reading or parsing {PULL_LIST_DB_FILE} for upcoming releases: {e}.",
                            color=0xFF0000,
                            is_dry_run=args.dry_run,
                        )
                
                # Check for next Wednesday releases even on non-Wednesday
                _check_next_wednesday_releases(all_comics_from_json_for_upcoming, args.dry_run)

            else:
                logger.error(
                    "Failed to update JSON pull list from downloaded Excel file."
                )
            
            try:
                os.remove(pulled_comics_source_file)
                logger.info(f"Cleaned up downloaded file: {pulled_comics_source_file}")
            except Exception as e:
                logger.warning(
                    f"Warning: Could not remove temporary file {pulled_comics_source_file}: {e}"
                )
        else:
            logger.error(
                "Failed to download pull list from LCG. Cannot update JSON file."
            )
        
        script_end_time = datetime.now()
        script_duration = script_end_time - script_start_time
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Comic Grabber Bot Status (Non-Wednesday)",
            description=(
                f"Script execution completed at: {script_end_time.strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Duration: {script_duration}"
            ),
            color=0x3498DB,  # Blue color for informational end
            is_dry_run=args.dry_run,
        )
        return # Exit if not Wednesday and not running with --excel-file

    # --- Logic for Wednesday (or if --excel-file was provided) ---
    check_date = datetime.now().date()
    if is_wednesday:
        logger.info(
            f"Today is Wednesday. Checking for comics released on: {check_date.strftime('%Y-%m-%d')}"
        )
    else:
        logger.info(
            f"Running with --excel-file. Processing and checking for comics as per argument."
        )


    # Determine source of pull list: downloaded file or fresh login/download
    pulled_comics_source_file = None
    file_was_downloaded = (
        False  # Flag to know if we should delete the temporary Excel file
    )

    if args.excel_file:
        if os.path.exists(args.excel_file):
            pulled_comics_source_file = args.excel_file
            logger.info(f"Using provided Excel file: {pulled_comics_source_file}")
            file_was_downloaded = False  # Not downloaded in this run
        else:
            logger.error(
                f"Error: Provided Excel file '{args.excel_file}' not found. Proceeding with fresh download."
            )
            send_discord_notification(
                webhook_url=DISCORD_WEBHOOK_URL,
                title="Warning: Excel File Not Found",
                description=f"Provided Excel file '{args.excel_file}' not found. Attempting fresh download.",
                color=0xFF8C00,
                is_dry_run=args.dry_run,
            )
            pulled_comics_source_file = login_and_download_pull_list()
            if pulled_comics_source_file:
                file_was_downloaded = True
    else:
        pulled_comics_source_file = login_and_download_pull_list()
        if pulled_comics_source_file:
            file_was_downloaded = True

    if not pulled_comics_source_file:
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Critical Error: LCG Pull List Failed",
            description="Could not obtain LCG pull list (no file provided or download failed). Exiting.",
            color=0xFF0000,
            is_dry_run=args.dry_run,
        )
        return

    # Process the Excel file and update the JSON pull list file
    json_update_success = update_json_pull_list_from_excel(pulled_comics_source_file)

    # Clean up the downloaded file if it was downloaded in this run
    if file_was_downloaded and os.path.exists(pulled_comics_source_file):
        try:
            os.remove(pulled_comics_source_file)
            logger.info(f"Cleaned up downloaded file: {pulled_comics_source_file}")
        except Exception as e:
            logger.warning(
                f"Warning: Could not remove temporary file {pulled_comics_source_file}: {e}"
            )

    # If --excel-file was provided, we've just updated the JSON, so exit.
    if args.excel_file and not is_wednesday: # Added check for is_wednesday
        logger.info(
            "--- Excel file provided on non-Wednesday. JSON pull list updated. Exiting without daily download check. ---"
        )
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="JSON Sync Complete (via --excel-file)",
            description="Excel file provided. JSON pull list updated.",
            color=0x3498DB,
            is_dry_run=args.dry_run,
        )
        return


    # --- Proceed with Daily Download Check (only if not in --excel-file mode and it IS Wednesday, or if --excel-file was used with search-past-releases) ---
    if not json_update_success:
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Error: JSON Update Failed",
            description="JSON pull list update from LCG failed. Skipping daily download check.",
            color=0xFF0000,
            is_dry_run=args.dry_run,
        )
        return

    # Load comics from the JSON file for today's release check
    all_comics_from_json = []
    if os.path.exists(PULL_LIST_DB_FILE):
        try:
            with open(PULL_LIST_DB_FILE, "r", encoding="utf-8") as f:
                all_comics_from_json = json.load(f)
        except (json.JSONDecodeError, FileNotFoundError) as e:
            logger.error(
                f"Error reading or parsing {PULL_LIST_DB_FILE} for releases: {e}. Skipping daily check."
            )
            send_discord_notification(
                webhook_url=DISCORD_WEBHOOK_URL,
                title="Error: JSON Read Failed",
                description=f"Error reading or parsing {PULL_LIST_DB_FILE} for releases: {e}. Skipping daily check.",
                color=0xFF0000,
                is_dry_run=args.dry_run,
            )
    else:
        logger.warning(
            f"JSON pull list file '{PULL_LIST_DB_FILE}' not found. No comics to check today."
        )
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Warning: JSON File Not Found",
            description=f"JSON pull list file '{PULL_LIST_DB_FILE}' not found. No comics to check today.",
            color=0xFFFF00,
            is_dry_run=args.dry_run,
        )

    comics_to_search = []
    today_date = datetime.now().date()
    for comic_data in all_comics_from_json:
        release_date_obj = datetime.strptime(
            comic_data.get("release_date"), "%Y-%m-%d"
        ).date()
        if args.search_past_releases:
            if release_date_obj <= today_date:
                comics_to_search.append(
                    {
                        "series_name": comic_data.get("comic_name"),
                        "issue_number": None,
                        "release_date": release_date_obj,
                    }
                )
        elif release_date_obj == today_date: # Only check today's releases if not searching past
            comics_to_search.append(
                {
                    "series_name": comic_data.get("comic_name"),
                    "issue_number": None,
                    "release_date": release_date_obj,
                }
            )

    if args.search_past_releases:
        logger.info(
            f"Found {len(comics_to_search)} comics in {PULL_LIST_DB_FILE} to search for (including past releases)."
        )
    else:
        logger.info(
            f"Found {len(comics_to_search)} comics released today in {PULL_LIST_DB_FILE}."
        )

    if not comics_to_search:
        if args.search_past_releases:
            logger.info(
                f"No comics found in your JSON file to search for (even including past releases)."
            )
            send_discord_notification(
                webhook_url=DISCORD_WEBHOOK_URL,
                title="No Comics Found (Past Releases)",
                description=f"No comics found in your JSON file to search for (even including past releases).",
                color=0xFFFF00,
                is_dry_run=args.dry_run,
            )
        else:
            logger.info(
                f"No comics from your pull list are released today ({check_date.strftime('%Y-%m-%d')})."
            )
            send_discord_notification(
                webhook_url=DISCORD_WEBHOOK_URL,
                title="No Comics for Today",
                description=f"No comics from your pull list are released today ({check_date.strftime('%Y-%m-%d')}).",
                color=0xFFFF00,
                is_dry_run=args.dry_run,
            )
        # Check for next Wednesday releases even if no comics are for today/past
        _check_next_wednesday_releases(all_comics_from_json, args.dry_run)
        return

    logger.info(f"Found {len(comics_to_search)} comics to process for download.")
    grabbed_count = 0
    skipped_count = 0 # Added skipped count
    failed_count = 0

    # Iterate through selected comics and attempt download via AirDC++
    for comic in comics_to_search:
        comic_title_full = f"{comic['series_name']}"
        logger.info(
            f"Processing comic: {comic_title_full} (Release Date: {comic['release_date'].strftime('%Y-%m-%d')})"
        )

        # search_airdcpp now returns a tuple: (found_match_info, session_search_id)
        found_match_info, session_id_for_search = search_airdcpp(
            comic_title_full, is_dry_run=args.dry_run
        )

        if found_match_info and session_id_for_search:  # Ensure both are returned
            # `found_match_info` now contains 'id', 'name', 'path', 'size', 'tth'
            # We no longer construct a `target_path_hint` here in the main loop,
            # as download_airdcpp will derive `target_name` from `found_match_info['name']`.

            download_result = download_airdcpp( # Changed variable name to download_result
                found_match_info, session_id_for_search, is_dry_run=args.dry_run
            )

            if download_result is True: # Check if it's explicitly True for success
                grabbed_count += 1
                notification_message = (
                    f"**Queued download for:** {comic_title_full}\n"
                    f"Found file: `{found_match_info['path']}`\n"
                    f"Size: {round(found_match_info['size'] / (1024*1024), 2)} MB\n"
                    f"Will download to AirDC++'s default folder as: `{found_match_info['name']}`"
                )
                send_discord_notification(
                    webhook_url=DISCORD_WEBHOOK_URL,
                    title="Comic Download Queued",
                    description=notification_message,
                    color=0x00FF00,
                    is_dry_run=args.dry_run,
                )
            elif download_result == "skipped": # Check if it's "skipped"
                skipped_count += 1
                # Notification is already sent within download_airdcpp for this case
            else: # Must be False for a general failure
                failed_count += 1
        else:
            failed_count += 1
            logger.info(
                f"No suitable file found on AirDC++ for {comic_title_full}. Skipping download."
            )
            # Notification for "No Download Found" is already sent within search_airdcpp if match is None
        time.sleep(1)

    # After processing today's/past releases, check for next Wednesday
    _check_next_wednesday_releases(all_comics_from_json, args.dry_run)

    script_end_time = datetime.now()  # Define script_end_time here
    script_duration = script_end_time - script_start_time

    logger.info(f"--- Comic Grabber Bot Finished ---")
    final_message = (
        f"**Daily Run Complete!**\n"
        f"Successfully queued: {grabbed_count} comic(s)\n"
        f"Skipped (already exists): {skipped_count} comic(s)\n" # Added skipped count to summary
        f"Failed to find/queue: {failed_count} comic(s)"
    )
    send_discord_notification(
        webhook_url=DISCORD_WEBHOOK_URL,
        title="Daily Run Summary",
        description=final_message,
        color=0x00FF00 if failed_count == 0 else 0xFF8C00,
        is_dry_run=args.dry_run,
    )
    # Send final script status notification
    send_discord_notification(
        webhook_url=DISCORD_WEBHOOK_URL,
        title="Comic Grabber Bot Status",
        description=(
            f"Script execution completed at: {script_end_time.strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Duration: {script_duration}"
        ),
        color=0x3498DB,  # Blue color for informational end
        is_dry_run=args.dry_run,
    )


def _check_next_wednesday_releases(all_comics, is_dry_run=False):
    """
    Checks for comics releasing next Wednesday and sends a Discord notification
    listing those comics.

    Args:
        all_comics (list): A list of dictionaries, where each dictionary represents a comic
                           and contains at least 'comic_name' and 'release_date'.
        is_dry_run (bool, optional): If True, indicates a dry run, affecting notifications. Defaults to False.
    """
    today = datetime.now().date()
    # Calculate next Wednesday
    # Weekday() returns 0 for Monday, 1 for Tuesday, ..., 6 for Sunday
    # We want 2 for Wednesday.
    days_until_next_wednesday = (2 - today.weekday() + 7) % 7
    if (
        days_until_next_wednesday == 0
    ):  # If today is Wednesday, it means the *next* Wednesday, so add 7 days.
        days_until_next_wednesday = 7

    next_wednesday = today + timedelta(days=days_until_next_wednesday)

    logger.info(
        f"Checking for comics releasing next Wednesday: {next_wednesday.strftime('%Y-%m-%d')}"
    )

    next_wednesday_comics = []
    for comic_data in all_comics:
        try:
            release_date_obj = datetime.strptime(
                comic_data.get("release_date"), "%Y-%m-%d"
            ).date()
            if release_date_obj == next_wednesday:
                next_wednesday_comics.append(comic_data["comic_name"])
        except ValueError:
            # Date parsing errors are handled during update_json_pull_list_from_excel
            pass

    if next_wednesday_comics:
        message = f"**Comics Releasing Next Wednesday ({next_wednesday.strftime('%Y-%m-%d')}):**\n"
        for comic_name in sorted(next_wednesday_comics):
            message += f"- {comic_name}\n"
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Upcoming Comic Releases",
            description=message,
            color=0x3498DB,
            is_dry_run=is_dry_run,
        )
        logger.info(
            f"Sent Discord notification for {len(next_wednesday_comics)} comics releasing next Wednesday."
        )
    else:
        logger.info(
            f"No comics found releasing next Wednesday ({next_wednesday.strftime('%Y-%m-%d')})."
        )
        send_discord_notification(
            webhook_url=DISCORD_WEBHOOK_URL,
            title="Upcoming Comic Releases",
            description=f"No comics scheduled for release next Wednesday ({next_wednesday.strftime('%Y-%m-%d')}).",
            color=0xADD8E6,
            is_dry_run=is_dry_run,
        )


if __name__ == "__main__":
    main()
