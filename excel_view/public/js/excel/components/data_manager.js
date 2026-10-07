/**
 * excel_view/components/data_manager.js
 *
 * Handles all Frappe DB interactions:
 *  - Convert list data → 2D matrix for HOT
 *  - Determine cell read-only state from permissions + field config
 *  - Debounced save of edited cells → full ORM doc.save() via bulk_set_value
 *  - Row deletion
 */

frappe.provide("frappe.views.excel");

frappe.views.excel.DataManager = class DataManager {
	/**
	 * @param {Object} opts
	 * @param {Object} opts.board - ExcelBoard instance
	 */
	constructor(opts) {
		this.board = opts.board;
		// { doc_name: { fieldname: value, ... }, ... }
		this._save_queue = {};
		this._save_debounced = frappe.utils.debounce(this._flush_saves.bind(this), 800);
		this._dirty = false;
	}

	// ── Data transformation ───────────────────────────────────────────────────

	/**
	 * Convert Frappe list data (array of objects) → 2D matrix aligned to HOT columns.
	 * @param {Object[]} data
	 * @param {Object[]} columns - HOT column config array
	 * @returns {Array[]}
	 */
	to_matrix(data, columns) {
		return data.map((row) =>
			columns.map((col) => {
				const val = row[col.data];
				// Normalise null/undefined to empty string
				return val !== null && val !== undefined ? val : "";
			})
		);
	}

	// ── Cell meta (HOT `cells` callback) ─────────────────────────────────────

	/**
	 * Called by HOT for every rendered cell.
	 * Returns { readOnly, className } based on permissions + field config.
	 * @param {number} row
	 * @param {number} col
	 * @returns {Object}
	 */
	get_cell_meta(row, col) {
		const col_def = this.board.columns[col];
		if (!col_def) return {};

		const is_readonly =
			col_def._readonly ||
			!this.board.list_view.can_write ||
			// name column always read-only
			col_def._is_name_col;

		if (is_readonly) {
			return { readOnly: true, className: "htDimmed" };
		}
		return {};
	}

	// ── Save queue ────────────────────────────────────────────────────────────

	/**
	 * Queue a batch of cell changes for debounced save to Frappe DB.
	 * @param {Array[]} changes - HOT changes [[row, col, oldVal, newVal], ...]
	 */
	queue_save(changes) {
		if (!changes || !this.board.list_view.can_write) return;

		const data = this.board.list_view.data;
		const columns = this.board.columns;
		let any_queued = false;

		changes.forEach(([row, fieldname, , newVal]) => {
			const doc = data[row];
			// In array-of-objects mode, HOT gives fieldname as 'col' (the prop key)
			const col_def = columns.find((c) => c.data === fieldname);
			// Skip formula/join columns — they exist only in the grid, never in the DB
			// Skip _is_new rows — they are pending inline inserts, not yet in the DB
			if (
				!doc ||
				doc._is_new ||
				!col_def ||
				col_def._readonly ||
				col_def._is_name_col ||
				col_def._is_formula_col ||
				col_def._is_join_col
			)
				return;

			const doc_name = doc.name;
			if (!doc_name || !fieldname) return;

			if (!this._save_queue[doc_name]) this._save_queue[doc_name] = {};
			this._save_queue[doc_name][fieldname] = newVal;
			any_queued = true;
		});

		// Nothing was actually queued (e.g. all changes were formula/readonly columns)
		// — don't show the dirty indicator or schedule a pointless flush.
		if (!any_queued) return;

		this._dirty = true;
		this._show_dirty_indicator();
		// Manual-save mode: edits wait until Ctrl+S (or leaving the page). Otherwise
		// they are saved about a second after the last edit.
		if (!this.board.manual_save) this._save_debounced();
	}

	/**
	 * Flush all queued saves to Frappe DB.
	 * Called automatically after debounce delay.
	 */
	async _flush_saves() {
		if (!Object.keys(this._save_queue).length) return;

		const queue = { ...this._save_queue };
		this._save_queue = {};

		const doctype = this.board.doctype;
		// Single HTTP round-trip: server calls doc.save() for each record,
		// running all controllers and hooks in one transaction.
		const updates = Object.entries(queue).map(([name, fields]) => ({ name, fields }));

		try {
			const r = await frappe.call({
				method: "excel_view.api.bulk_set_value",
				args: { doctype, updates: JSON.stringify(updates) },
				// Suppress Frappe's default msgprint on error — we show show_error instead
				error: () => {},
			});
			const errors = r.message?.errors || [];
			this._mark_failed_cells(errors, updates);
			if (errors.length) {
				// Show the first error in the friendly dialog; toast a count if multiple.
				// Every failed cell is also marked red (see _mark_failed_cells).
				const first = errors[0];
				frappe.views.excel.show_error(
					{
						_server_messages: JSON.stringify([
							JSON.stringify({ message: first.error }),
						]),
					},
					errors.length > 1
						? __("{0} record(s) could not be saved", [errors.length])
						: __("Saving {0}", [first.name])
				);
			} else {
				this._dirty = false;
				this._clear_dirty_indicator();
				frappe.views.excel.toast(__("Saved"), "success", 2000);
			}
		} catch (err) {
			frappe.views.excel.show_error(err, __("Saving changes to {0}", [doctype]));
		}
	}

	// ── Row deletion ──────────────────────────────────────────────────────────

	/**
	 * Delete rows from HOT and from Frappe DB.
	 * @param {number} start_row - Inclusive start row index
	 * @param {number} end_row   - Inclusive end row index
	 */
	async delete_rows(start_row, end_row) {
		const data = this.board.list_view.data;
		const to_delete = data.slice(start_row, end_row + 1);

		const failures = [];
		for (const doc of to_delete) {
			try {
				await frappe.db.delete_doc(this.board.doctype, doc.name);
			} catch (e) {
				failures.push(doc.name);
				console.error("Excel View delete error:", e);
			}
		}

		if (failures.length) {
			frappe.show_alert(
				{
					message: __("Failed to delete: {0}", [failures.join(", ")]),
					indicator: "red",
				},
				5
			);
		} else {
			frappe.show_alert(
				{
					message: __("{0} record(s) deleted", [to_delete.length]),
					indicator: "green",
				},
				2
			);
			// Remove from HOT and refresh
			this.board.hot.alter("remove_row", start_row, end_row - start_row + 1);
			this.board.list_view.refresh();
		}
	}

	// ── Dirty state ───────────────────────────────────────────────────────────

	_show_dirty_indicator() {
		this.board.list_view?.page?.set_indicator(__("Unsaved changes"), "orange");
	}

	_clear_dirty_indicator() {
		this.board.list_view?.page?.clear_indicator();
	}

	/**
	 * Returns true if there are unsaved changes.
	 */
	is_dirty() {
		return this._dirty || Object.keys(this._save_queue).length > 0;
	}

	// ── Failed rows ───────────────────────────────────────────────────────────

	/**
	 * Colour every cell the server refused, with the server's message as a tooltip.
	 * The server returns one entry per failed row: {name, error, fields: [fieldnames]}.
	 * A failed cell stays marked until the user edits it again.
	 */
	_mark_failed_cells(errors, updates) {
		const board = this.board;
		if (!board._failed_cells) board._failed_cells = new Map();
		const plain = (s) =>
			String(s || "")
				.replace(/<[^>]*>/g, "")
				.trim();

		// A successful save of a cell clears an older mark on it.
		const failed_rows = new Set(errors.map((e) => e.name));
		updates.forEach((u) => {
			if (failed_rows.has(u.name)) return;
			Object.keys(u.fields).forEach((f) => board._failed_cells.delete(`${u.name}|${f}`));
		});

		errors.forEach((e) => {
			const upd = updates.find((u) => u.name === e.name);
			const fields = e.fields || Object.keys(upd?.fields || {});
			fields.forEach((f) => board._failed_cells.set(`${e.name}|${f}`, plain(e.error)));
		});
		board.hot?.render();

		if (errors.length) {
			// The server did not store these edits: reload so the cells show what is saved.
			setTimeout(() => board.list_view?.refresh(), 300);
		}
	}
};
