# =============================================================================
# Spine HR Configuration
# Update these settings before running the dashboard
# =============================================================================

SPINE_URL = "https://inovatix.spinehrm.in"
USERNAME = "IIL56"
PASSWORD = "PassworD@1993"
LOGIN_FOR = "Employee"   # Change to "Admin" if you login as Admin

# How many months back to fetch (0 = current month only)
MONTHS_BACK = 0

# Chrome driver settings
HEADLESS = False         # Set False to see the browser while scraping
CHROME_TIMEOUT = 30      # seconds to wait for page elements

# Data output
DATA_FILE = "essl_attendance.json"   # eSSL biometric machine data
