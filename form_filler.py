"""
Form filler for ES Windows order form.
Uses native Selenium Select for <select> elements and standard input handling.
Fields enable cascadingly — each selection enables the next field.
If a field stays disabled or a value doesn't exist in a dropdown, the process terminates.
"""

import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    ElementNotInteractableException,
)
import form_selectors as sel


class FormFillerError(Exception):
    """Raised when a field cannot be filled — terminates the process."""

    pass


class FormFiller:
    def __init__(self, driver, wait_timeout=15):
        self.driver = driver
        self.wait_timeout = wait_timeout

    # ─── Main Entry Point ────────────────────────────────────

    def add_line_item(self, row_data):
        """Click Add Line Item -> fill form -> click Create."""
        self._click_add_line_item()
        self._wait_for_modal()
        self._fill_form(row_data)
        self._click_create()
        self._wait_for_modal_close()

    # ─── Flow Steps ──────────────────────────────────────────

    def _click_add_line_item(self):
        btn = self._find_visible(sel.ADD_LINE_ITEM_BTN)
        self.driver.execute_script("arguments[0].scrollIntoView(true);", btn)
        time.sleep(0.5)
        btn.click()
        time.sleep(2)

    def _wait_for_modal(self):
        """Wait for the modal to become visible."""
        WebDriverWait(self.driver, self.wait_timeout).until(
            lambda d: self._get_active_modal() is not None
        )
        time.sleep(1)

    def _click_create(self):
        btn = self._find_visible(sel.SUBMIT_BTN)
        self.driver.execute_script("arguments[0].scrollIntoView(true);", btn)
        time.sleep(0.5)
        btn.click()
        time.sleep(5)

    def _wait_for_modal_close(self):
        for _ in range(20):
            if self._get_active_modal() is None:
                return
            time.sleep(1)
        time.sleep(2)

    # ─── Modal Helpers ───────────────────────────────────────

    def _get_active_modal(self):
        modals = self.driver.find_elements(By.CSS_SELECTOR, sel.MODAL)
        for m in modals:
            style = m.get_attribute("style") or ""
            if "display:none" not in style.replace(" ", "").lower():
                return m
        return None

    def _find_visible(self, css_selector):
        elements = self.driver.find_elements(By.CSS_SELECTOR, css_selector)
        for el in elements:
            try:
                if el.is_displayed():
                    return el
            except Exception:
                continue
        if elements:
            return elements[-1]
        raise NoSuchElementException(f"No element found: {css_selector}")

    # ─── Form Filling (left to right, matching Excel column order) ────

    def _fill_form(self, data):
        """Fill fields in Excel column order: A through R."""

        # A: Product Type
        self._select_dropdown(
            sel.PRODUCT_TYPE, data.get("Product Type"), "Product Type"
        )

        # B: Brand
        self._select_dropdown(sel.BRAND, data.get("Brand"), "Brand")

        # C: Category
        self._select_dropdown(sel.CATEGORY, data.get("Category"), "Category")

        # D: Rating
        self._select_dropdown(sel.RATING, data.get("Rating"), "Rating")

        # E: Model
        self._select_dropdown(sel.MODEL, data.get("Model"), "Model")

        # F: Configuration
        if data.get("Configuration"):
            self._select_dropdown(
                sel.CONFIGURATION, data["Configuration"], "Configuration"
            )

        # G: Max External PSF (optional)
        if data.get("Max External PSF"):
            self._fill_input(
                sel.MAX_EXT_PSF, data["Max External PSF"], "Max External PSF"
            )

        # H: StoreFront Door — always select it (even NONE) to enable downstream fields
        storefront_door = data.get("StoreFront Door")
        if storefront_door:
            self._select_dropdown(
                sel.STOREFRONT_DOOR, storefront_door, "StoreFront Door"
            )

            # I: Door Width (only when door is not NONE)
            if str(storefront_door).strip().upper() != "NONE" and data.get(
                "Door Width"
            ):
                self._fill_input(
                    sel.STOREFRONT_DOOR_WIDTH, data["Door Width"], "Door Width"
                )

        # J: Width
        if data.get("Width"):
            self._fill_input(sel.WIDTH, data["Width"], "Width")

        # K: Panels
        if data.get("Panels"):
            self._select_dropdown(
                sel.STOREFRONT_PANELS, str(int(data["Panels"])), "Panels"
            )

        # L: Door Panels (only when door is not NONE)
        if (
            storefront_door
            and str(storefront_door).strip().upper() != "NONE"
            and data.get("Door Panels")
        ):
            self._select_dropdown(
                sel.STOREFRONT_DOOR_PANEL, str(int(data["Door Panels"])), "Door Panels"
            )

        # M: Height
        if data.get("Height"):
            self._fill_input(sel.HEIGHT, data["Height"], "Height")

        # N: Aluminum Finish — skip (inherited from order header)

        # O: Glass Type
        if data.get("Glass Type"):
            self._select_dropdown(sel.GLASS_TYPE, data["Glass Type"], "Glass Type")

        # P: Glass Color — skip (inherited from order header)

        # Q: LOW-E checkbox
        if str(data.get("LOW-E", "")).strip().lower() == "yes":
            self._set_checkbox(sel.LOW_E, True, "LOW-E")

        # R: Privacy checkbox
        if str(data.get("Privacy", "")).strip().lower() == "yes":
            self._set_checkbox(sel.PRIVACY, True, "Privacy")

    # ─── Wait for field to be enabled ────────────────────────

    def _wait_for_enabled(self, css_selector, field_name):
        """Wait for a field to become enabled. Terminate if it stays disabled."""
        try:
            WebDriverWait(self.driver, self.wait_timeout).until(
                lambda d: self._is_field_enabled(css_selector)
            )
        except TimeoutException:
            raise FormFillerError(
                f"TERMINATED: '{field_name}' is still disabled after {self.wait_timeout}s. "
                f"A previous field may have an invalid value."
            )
        time.sleep(0.5)

    def _is_field_enabled(self, css_selector):
        elements = self.driver.find_elements(By.CSS_SELECTOR, css_selector)
        for el in reversed(elements):
            try:
                if (
                    el.is_displayed()
                    or el.get_attribute("style") is None
                    or "display:none" not in (el.get_attribute("style") or "")
                ):
                    return not el.get_attribute("disabled")
            except Exception:
                continue
        return False

    # ─── Dropdown Selection ──────────────────────────────────

    def _select_dropdown(self, css_selector, value, field_name):
        """Wait for dropdown to enable, then select value by visible text."""
        if not value:
            return

        value = str(value).strip()
        print(f"    {field_name} = {value}")

        # Wait for the field to be enabled
        self._wait_for_enabled(css_selector, field_name)

        # Find the active (visible) select element
        el = self._find_visible(css_selector)

        # Use Selenium's Select to pick the option by visible text
        select = Select(el)
        available = [o.text.strip() for o in select.options if o.text.strip()]

        # Try exact match first
        for option in select.options:
            if option.text.strip().upper() == value.upper():
                select.select_by_visible_text(option.text.strip())
                time.sleep(2)
                return

        # Try partial match (value contained in option text)
        for option in select.options:
            if value.upper() in option.text.strip().upper():
                select.select_by_visible_text(option.text.strip())
                time.sleep(2)
                return

        # Value not found — terminate
        raise FormFillerError(
            f"TERMINATED: Value '{value}' not found in '{field_name}' dropdown. "
            f"Available options: {available}"
        )

    # ─── Text Input ──────────────────────────────────────────

    def _fill_input(self, css_selector, value, field_name):
        """Wait for input to enable, then type the value."""
        if not value:
            return

        value = str(value).strip()
        print(f"    {field_name} = {value}")

        self._wait_for_enabled(css_selector, field_name)

        el = self._find_visible(css_selector)
        el.clear()
        el.send_keys(value)
        time.sleep(0.5)

        # Click away from the field to trigger blur event
        # This enables the next field via the site's JS
        modal = self._get_active_modal()
        if modal:
            # Click the modal title (safe neutral area)
            try:
                title = modal.find_element(By.CSS_SELECTOR, "h4")
                title.click()
            except Exception:
                self.driver.execute_script("arguments[0].blur();", el)
        else:
            self.driver.execute_script("arguments[0].blur();", el)
        time.sleep(2)

    # ─── Checkbox ────────────────────────────────────────────

    def _set_checkbox(self, css_selector, checked, field_name):
        """Wait for checkbox to enable, then click if needed."""
        print(f"    {field_name} = {'Yes' if checked else 'No'}")

        self._wait_for_enabled(css_selector, field_name)

        el = self._find_visible(css_selector)
        if el.is_selected() != checked:
            el.click()
        time.sleep(0.5)
