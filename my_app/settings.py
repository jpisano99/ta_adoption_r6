from datetime import datetime
from my_app.my_secrets import passwords
import os

# database configuration settings
database = dict(
    DATABASE="cust_ref_db",
    USER="root",
    PASSWORD=passwords["DB_PASSWORD"],
    HOST="localhost"
)

# Smart sheet Config settings
ss_token = dict(
    SS_TOKEN=passwords["SS_TOKEN"]
)

# application predefined constants
app_cfg = dict(
    VERSION=1.0,
    GITHUB="{url}",
    HOME=os.path.expanduser("~"),
    WORKING_DIR='ta_adoption_data',
    UPDATES_DIR='ta_data_updates',
    ARCHIVES_DIR='archives',
    PROD_DATE='',
    UPDATE_DATE='',
    META_DATA_FILE='config_data.json',

    RAW_SUBSCRIPTIONS='TA Master Subscriptions as of',
    RAW_BOOKINGS='TA Master Bookings with SO as of',
    RAW_TA_AS_FIXED_SKU='TA AS Delivery Status as of',

    XLS_SUBSCRIPTIONS='tmp_Master Subscriptions.xlsx',
    XLS_AS_DELIVERY_STATUS='tmp_AS Delivery.xlsx',
    XLS_BOOKINGS='tmp_TA Master Bookings.xlsx',
    XLS_AS_SKUS='tmp_TA AS SKUs.xlsx',
    XLS_CUSTOMER='tmp_TA Customer List.xlsx',
    XLS_DASHBOARD='tmp_TA Unified Adoption Dashboard.xlsx',
    SS_SAAS='SaaS customer tracking',
    SS_CX='CX Tetration Customer Comments v3.0',
    # SS_CX='Tetration Engaged Customer Report',
    SS_AS='Tetration Shipping Notification & Invoicing Status',
    SS_COVERAGE='Tetration Coverage Map',
    SS_SKU='Tetration SKUs',
    SS_CUSTOMERS='TA Customer List',
    SS_DASHBOARD='TA Unified Adoption Dashboard',
    SS_WORKSPACE='Tetration Customer Adoption Workspace',
    AS_OF_DATE=datetime.now().strftime('_as_of_%m_%d_%Y')
)



