Comic Grabber Bot

This Python script automates the management of your comic pull list from League of Comic Geeks (LCG) and facilitates the download of released comics via an AirDC++ client. It's designed to keep your comic collection up-to-date with minimal manual intervention.
Table of Contents

    Features

    Prerequisites

    Installation

    Configuration

    Usage

    Logging

    Discord Notifications

    Troubleshooting & Notes

Features

    League of Comic Geeks (LCG) Integration: Automatically logs into your LCG account and downloads your latest pull list in Excel format.

    Local Pull List Management: Parses the downloaded Excel file and updates a local JSON database (pull_list.json). This database stores only future and current comic releases, effectively cleaning out old entries.

    AirDC++ Search and Download: For comics released on the current day (or specified past dates via command-line argument), the bot searches for matching files on AirDC++ hubs and queues them for download. It intelligently skips downloads for files that AirDC++ indicates already exist on disk or in its queue.

    Discord Notifications: Sends detailed, customizable notifications to a configured Discord webhook for various events, including:

        Script start and completion status.

        Successfully queued comic downloads.

        Skipped downloads (if the file already exists).

        Failed search or download attempts.

        Upcoming comic releases for the next Wednesday (new comic day).

    Robust Logging: Implements a comprehensive logging system that outputs essential information to the console and detailed debug information to timestamped log files. Old log files are automatically cleaned up.

    Command-Line Arguments: Supports flexible operation modes, allowing you to:

        Process a specific Excel file (e.g., for initial setup or manual updates).

        Search for all comics currently in your JSON pull list (including past releases).

        Perform "dry runs" to simulate downloads without actually initiating them.

Prerequisites

Before running the Comic Grabber Bot, ensure you have the following installed:

    Python 3.6+: The script is developed using modern Python features.

    AirDC++ Web API: Your AirDC++ client must have its Web API enabled and accessible from where the script is running. Ensure the API URL, username, and password are correct.

    League of Comic Geeks Account: You need an active LCG account with a pull list configured.

Installation

    Clone the Repository (or download the script):

    git clone https://github.com/your-repo/comic-grabber-bot.git
    cd comic-grabber-bot


    (Replace https://github.com/your-repo/comic-grabber-bot.git with the actual repository URL if this script is hosted.)

    Create a Virtual Environment (Recommended):

    python3 -m venv venv
    source venv/bin/activate  # On Windows: `venv\Scripts\activate`


    Install Dependencies:

    pip install -r requirements.txt


    If requirements.txt is not provided, you'll need to install them manually:

    pip install requests pandas xlrd beautifulsoup4 discord-webhook python-dotenv


Configuration

The script uses environment variables for sensitive information and configuration. Create a file named .env in the same directory as comicgrabr.py.

.env File Example:

# League of Comic Geeks Credentials
LCG_USERNAME="your_lcg_username"
LCG_PASSWORD="your_lcg_password"

# AirDC++ Web API Details
# Ensure AirDC++ Web API is enabled and accessible (e.g., http://127.00.1:5600/api/v1/)
AIRDCPP_API_URL="http://unraid:5600/api/v1/"
AIRDCPP_USERNAME="your_airdcpp_api_username"
AIRDCPP_PASSWORD="your_airdcpp_api_password"

# Discord Webhook URL for Notifications (Optional)
# Create a webhook in your Discord server settings (Server Settings -> Integrations -> Webhooks)
DISCORD_WEBHOOK_URL="https://discord.com/api/webhooks/YOUR_WEBHOOK_ID/YOUR_WEBHOOK_TOKEN"

# Optional: qBittorrent details (if you have other scripts that use these, kept for consistency)
# DEFAULT_QB_HOST="http://localhost:8080"
# DEFAULT_QB_USER="admin"
# DEFAULT_QB_PASSWORD="adminadmin"


Important Notes:

    Replace placeholder values (your_lcg_username, etc.) with your actual credentials.

    Ensure AIRDCPP_API_URL points to the correct address and port of your AirDC++ Web API. If running on the same machine, http://127.0.0.1:5600/api/v1/ is common. If on another machine (like your Unraid server), use its IP or hostname (e.g., http://unraid:5600/api/v1/).

Usage

Navigate to the script's directory in your terminal and activate your virtual environment (if used).

cd /path/to/comic-grabber-bot
source venv/bin/activate # Or `venv\Scripts\activate` on Windows


Then run the script with Python:

python comicgrabr.py [OPTIONS]


Command-Line Options:

    --excel-file <path/to/excel.xls>, -f <path/to/excel.xls>

        Purpose: Use a specific, pre-downloaded LCG pull list Excel file instead of logging in and downloading a new one.

        Behavior: When this option is used, the script will parse the provided Excel file to update pull_list.json and then exit. It will NOT perform daily AirDC++ searches for released comics. This is useful for initial setup or manually updating your pull list.

        Example: python comicgrabr.py --excel-file ~/Downloads/my_pulls.xls

    --search-past-releases

        Purpose: Instructs the bot to search for all comics currently listed in pull_list.json that have a release date of today or in the past.

        Behavior: Typically, the script only searches for comics released on the current day (Wednesday, new comic day). This flag allows you to catch up on any missed downloads or for initial bulk grabbing of your past pull list.

        Example: python comicgrabr.py --search-past-releases

    --dry-run

        Purpose: Perform a simulation of the script's operations without making any actual changes (no downloads, no permanent state changes).

        Behavior: Discord notifications will be sent with a [DRY RUN] prefix, indicating no live actions were taken. This is excellent for testing your configuration.

        Example: python comicgrabr.py --dry-run

Typical Usage (Scheduled Task)

For daily automation, you would typically schedule this script to run. On Linux systems (like Unraid with a cron job), you might set it up to run nightly or weekly.

Example Cron Job (runs every Wednesday at 03:00 AM):

0 3 * * WED /path/to/comic-grabber-bot/venv/bin/python /path/to/comic-grabber-bot/comicgrabr.py >> /path/to/comic-grabber-bot/cron_output.log 2>&1


(Adjust /path/to/comic-grabber-bot/ to your actual script location.)

If you want it to run daily to update the pull list AND search for new comics on Wednesday, you might have two entries or run it daily:

Example Daily Cron Job (runs daily at 03:00 AM, handles Wednesday downloads automatically):

0 3 * * * /path/to/comic-grabber-bot/venv/bin/python /path/to/comic-grabber-bot/comicgrabr.py >> /path/to/comic-grabber-bot/cron_output.log 2>&1


Logging

The script logs its activity to both the console (standard output) and a dedicated log file within a logs/ directory in the script's root folder.

    Console Output: Provides a high-level overview (INFO and above).

    Log Files: Detailed debug information (DEBUG and above) is written to timestamped files (e.g., comic_grabber_bot_YYYY-MM-DD_HH-MM-SS.log).

    Log Retention: Log files older than LOG_RETENTION_DAYS (default 7 days, configurable within comicgrabr.py) are automatically cleaned up at the start of each run.

Discord Notifications

The bot sends rich embed messages to your configured Discord webhook for important events:

    Script Status: Start and end messages, including total duration.

    Download Status: Notifications for successfully queued downloads, explicitly skipped downloads (if the file already exists), and failed attempts.

    Configuration Errors: Alerts for missing credentials, API failures, or issues with Excel file parsing.

    Upcoming Releases: A summary of comics scheduled to release on the next Wednesday, providing a useful heads-up.

Troubleshooting & Notes

    AirDC++ API Accessibility: Ensure your AirDC++ Web API is running and accessible from the machine executing the script. Check firewall settings.

    LCG Credentials: If login fails, double-check your LCG username and password in the .env file. The website structure for login might also change, requiring script updates if CSRF token extraction fails.

    Excel File Format: The script specifically uses xlrd for .xls (Excel 97-2003) files. League of Comic Geeks exports in this format. If this changes, pandas.read_excel might need a different engine or parameters.

    Comic Naming: The script removes # and : from comic names extracted from Excel to improve search compatibility with AirDC++. If you find searches are consistently failing, examine the comic_name variable in the debug logs.

    Duplicate Downloads: The script now intelligently detects and skips download attempts if AirDC++ indicates the file already exists on disk or in its queue, preventing "400 Bad Request" errors.

    Timeouts: Increase timeout values in requests calls if you frequently experience timeout errors, especially over slower network connections.

    Docker/Containerization: For Unraid or other containerized environments, ensure that environment variables are correctly passed to the container running this script and that network access to AirDC++ is configured.
