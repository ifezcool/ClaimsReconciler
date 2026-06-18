import os
import logging
from dotenv import load_dotenv

load_dotenv('secrets.env')

NOTIFY_TO = os.getenv("NOTIFY_TO", "ifeoluwa.adeniyi@avonhealthcare.com")
NOTIFY_CC_CLAIMS = os.getenv("NOTIFY_CC_CLAIMS",
    "ifeoluwa.adeniyi@avonhealthcare.com,adedamola.ayeni@avonhealthcare.com,adebola.adesoyin@avonhealthcare.com,claims_officers@avonhealthcare.com,bi_dataanalytics@avonhealthcare.com,financedepartment@avonhealthcare.com")
NOTIFY_CC_FINANCE = os.getenv("NOTIFY_CC_FINANCE",
    "ifeoluwa.adeniyi@avonhealthcare.com,adedamola.ayeni@avonhealthcare.com,adebola.adesoyin@avonhealthcare.com,claims_officers@avonhealthcare.com,bi_dataanalytics@avonhealthcare.com,financedepartment@avonhealthcare.com")
NOTIFY_CC_DATE_VALIDATION = os.getenv("NOTIFY_CC_DATE_VALIDATION",
    "ifeoluwa.adeniyi@avonhealthcare.com,claims_officers@avonhealthcare.com,bi_dataanalytics@avonhealthcare.com")

CC_LISTS = {
    "default": NOTIFY_CC_CLAIMS,
    "claims": NOTIFY_CC_CLAIMS,
    "finance": NOTIFY_CC_FINANCE,
    "date_validation": NOTIFY_CC_DATE_VALIDATION,
}

def get_cc_list(key="default"):
    raw = CC_LISTS.get(key, CC_LISTS["default"])
    return [e.strip() for e in raw.split(",") if e.strip()]

def get_to_email():
    return NOTIFY_TO

LOG_FILE_ENABLED = os.getenv("LOG_FILE_ENABLED", "true").lower() in ("true", "1", "yes")
_handlers = [logging.StreamHandler()]
if LOG_FILE_ENABLED:
    try:
        _handlers.append(logging.FileHandler("claims_reconciler.log"))
    except (OSError, PermissionError):
        _handlers.append(logging.StreamHandler())

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=_handlers,
)
logger = logging.getLogger(__name__)
