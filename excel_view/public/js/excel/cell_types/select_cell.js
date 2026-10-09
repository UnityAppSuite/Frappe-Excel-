/**
 * excel_view/cell_types/select_cell.js
 *
 * HOT 6.x custom cell type for Frappe Select fields.
 * Renders the selected label.
 * Editor is a <select size="N"> listbox (always-visible options — no click-to-open needed)
 * positioned below the cell, styled like the Link editor dropdown.
 *
 * HOT 6.x BaseEditor contract:
 *   init()    – called once; create element and attach to DOM
 *   open()    – position + show element, load options
 *   close()   – hide element
 *   focus()   – MUST be defined; BaseEditor.beginEditing() calls it after open()
 *   getValue() / setValue()
 */

import Handsontable from "handsontable";

frappe.provide("frappe.views.excel");

// ── Renderer ─────────────────────────────────────────────────────────────────

function selectRenderer(hotInstance, td, row, col, prop, value, cellProperties) {
	Handsontable.renderers.TextRenderer.apply(this, arguments);
	if (value !== null && value !== undefined && value !== "") {
		td.innerText = __(String(value));
	}
	if (cellProperties.readOnly) td.classList.add("htDimmed");
}

// ── Editor ────────────────────────────────────────────────────────────────────

class SelectEditor extends Handsontable.editors.BaseEditor {
	/**
	 * Called once per HOT instance. Create the <select> listbox and attach to body.
	 * Using size > 1 makes the options always visible (no click-to-expand needed).
	 */
	init() {
		this._select = document.createElement("select");
		this._select.className = "ev-select-editor";
		Object.assign(this._select.style, {
			position: "fixed",
			display: "none",
			zIndex: "10000",
			border: "2px solid #1a73e8",
			borderRadius: "4px",
			padding: "0",
			fontSize: "13px",
			background: "var(--bg-color, #fff)",
			color: "var(--text-color, #333)",
			cursor: "pointer",
			boxSizing: "border-box",
			boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
			overflowY: "auto",
			minWidth: "120px",
		});

		// Pressing the mouse on the list must not reach the grid: it reads that as a click
		// outside the cell, closes the editor and saves the OLD value before "change" fires.
		this._select.addEventListener("mousedown", (e) => e.stopPropagation());

		// Click on an option → commit immediately
		this._select.addEventListener("change", () => this.finishEditing());

		this._select.addEventListener("keydown", (e) => {
			if (e.key === "Escape") {
				e.stopPropagation();
				this.finishEditing(true);
			}
			if (e.key === "Enter") {
				e.preventDefault();
				e.stopPropagation();
				this.finishEditing();
			}
		});

		this._select.addEventListener("blur", () => {
			requestAnimationFrame(() => {
				if (document.activeElement !== this._select) {
					this.finishEditing();
				}
			});
		});

		document.body.appendChild(this._select);
	}

	getValue() {
		return this._select.value;
	}

	setValue(newValue) {
		this._select.value = newValue != null ? String(newValue) : "";
	}

	/**
	 * open() – populate options as a visible listbox, position below the TD.
	 * this.cellProperties and this.TD are set by BaseEditor.prepare().
	 */
	open() {
		const source = this.cellProperties.source || [];

		// Build options — include blank "(none)" as first option
		this._select.innerHTML =
			`<option value="">${__("(none)")}</option>` +
			source
				.map(
					(opt) =>
						`<option value="${frappe.utils.escape_html(
							opt
						)}">${frappe.utils.escape_html(__(opt))}</option>`
				)
				.join("");

		// Set current value
		this._select.value = this.originalValue != null ? String(this.originalValue) : "";

		// Show as listbox: size = number of options (capped at 8)
		const totalOptions = source.length + 1; // +1 for (none)
		this._select.size = Math.min(totalOptions, 8);

		// Position below the cell (like Link editor dropdown)
		const rect = this.TD.getBoundingClientRect();
		const optionHeight = 24; // px per option row
		const listHeight = this._select.size * optionHeight;
		const viewportHeight = window.innerHeight;

		// Prefer below; flip above if not enough space
		let top;
		if (rect.bottom + listHeight + 4 < viewportHeight) {
			top = rect.bottom + 1;
		} else {
			top = rect.top - listHeight - 1;
		}

		Object.assign(this._select.style, {
			top: top + "px",
			left: rect.left + "px",
			width: Math.max(rect.width, 140) + "px",
			height: listHeight + "px",
			display: "block",
		});
	}

	close() {
		this._select.style.display = "none";
	}

	/**
	 * focus() MUST be defined — BaseEditor.beginEditing() calls it after open().
	 */
	focus() {
		this._select.focus();
	}

	prepare(row, col, prop, td, originalValue, cellProperties) {
		super.prepare(row, col, prop, td, originalValue, cellProperties);
	}
}

// ── Register ──────────────────────────────────────────────────────────────────

Handsontable.renderers.registerRenderer("ev-select", selectRenderer);
Handsontable.editors.registerEditor("ev-select", SelectEditor);
Handsontable.cellTypes.registerCellType("ev-select", {
	renderer: selectRenderer,
	editor: SelectEditor,
});
