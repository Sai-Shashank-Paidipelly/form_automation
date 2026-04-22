"""
CSS selectors for the ES Windows order form.
All fields use data-line-item-wizard-* attributes with native <select> and <input> elements.
Fields start disabled and enable cascadingly via JavaScript as you fill previous fields.
"""

ORDER_URL = "https://orders.eswindows.co/customer/sales_documents/v2/{order_number}/edit"

# ─── Main Page ───────────────────────────────────────────────
ADD_LINE_ITEM_BTN = "button.addLineItemButton"

# ─── Modal ───────────────────────────────────────────────────
MODAL = "div#new-item-modal"
SUBMIT_BTN = "button[data-line-item-wizard-submit_button]"
CANCEL_BTN = "a[data-line-item-wizard-cancel_button]"

# ─── Dropdown Fields (<select> elements) ─────────────────────
PRODUCT_TYPE = "select[data-line-item-wizard-product_type]"
BRAND = "select[data-line-item-wizard-brand]"
CATEGORY = "select[data-line-item-wizard-category]"
RATING = "select[data-line-item-wizard-impact_rating]"
MODEL = "select[data-line-item-wizard-assembly_system]"
CONFIGURATION = "select[data-line-item-wizard-configuration]"

# Storefront-specific
STOREFRONT_DOOR = "select[data-line-item-wizard-storefront_door]"
STOREFRONT_PANELS = "select[data-line-item-wizard-storefront_panels]"
STOREFRONT_DOOR_PANEL = "select[data-line-item-wizard-storefront_door_panel]"

# Finish & Glass
ALUMINUM_FINISH = "select[data-line-item-wizard-aluminum_finish]"
GLASS_TYPE = "select[data-line-item-wizard-glass_type]"
GLASS_COLOR = "select[data-line-item-wizard-glass_color]"

# ─── Text Input Fields ───────────────────────────────────────
MAX_EXT_PSF = "input[data-line-item-wizard-max_ext_psf]"
MAX_INT_PSF = "input[data-line-item-wizard-max_int_psf]"
STOREFRONT_DOOR_WIDTH = "input[data-line-item-wizard-storefront_door_target_width]"
WIDTH = "input[data-line-item-wizard-target_width]"
HEIGHT = "input[data-line-item-wizard-target_height]"
LINE_ITEM_NAME = "input[data-line-item-wizard-name]"

# ─── Checkboxes ──────────────────────────────────────────────
LOW_E = "input[data-line-item-wizard-low_e]"
PRIVACY = "input[data-line-item-wizard-privacy]"
