/**
 * excel_view/components/undo_manager.js
 *
 * Keeps Undo / Redo working after the grid reloads its rows.
 *
 * Why: Handsontable's built-in undo list is erased by every hot.loadData(). The grid
 * auto-saves about 1 second after an edit, Frappe then announces list_update, and the
 * grid reloads all rows (~1.5 s after the edit). Undo was therefore only available for
 * about 1.5 seconds.
 *
 * How: this manager keeps its own list. Each entry is keyed by record NAME and FIELD
 * (not by row number), so it still finds the right cell after a reload re-sorts the
 * rows. Undo writes the old value back through setDataAtCell, so the normal save path
 * stores it on the server.
 *
 * Entries are tied to the sheet they were made on, and are skipped when the sheet is
 * not the active one or the record is no longer in the grid.
 */

frappe.provide("frappe.views.excel");

frappe.views.excel.UndoManager = class UndoManager {
	/** @param {Object} opts  @param {Object} opts.board - ExcelBoard instance */
	constructor(opts) {
		this.board = opts.board;
		this._undo = [];
		this._redo = [];
		this._max = 100;
	}

	// Must be called once, right after the Handsontable instance exists.
	attach(hot) {
		this.hot = hot;
		const plugin = hot.undoRedo;
		if (!plugin) return;

		hot.addHook("afterChange", (changes, source) => this._record(changes, source));

		plugin.undo = () => this._step(this._undo, this._redo, true, "ev_undo");
		plugin.redo = () => this._step(this._redo, this._undo, false, "ev_redo");
		plugin.isUndoAvailable = () => this._undo.length > 0;
		plugin.isRedoAvailable = () => this._redo.length > 0;
		// Handsontable calls clear() on every loadData(); that is exactly what must not
		// wipe the list. Entries for rows that are gone are skipped when applied.
		plugin.clear = () => {};
	}

	/** Forget everything, for example when a different workbook is loaded. */
	reset() {
		this._undo.length = 0;
		this._redo.length = 0;
	}

	_sheet_id() {
		return this.board.sheet_manager?.get_current?.()?.id || null;
	}

	_record(changes, source) {
		if (!changes || source === "loadData" || source === "ev_undo" || source === "ev_redo")
			return;
		const data = this.board.list_view?.data;
		if (!data) return;
		const sheet = this._sheet_id();
		const group = [];
		changes.forEach(([row, prop, old_v, new_v]) => {
			const rec = data[row];
			if (!rec || typeof prop !== "string" || old_v === new_v) return;
			group.push({ name: rec.name, prop, old_v, new_v, sheet });
		});
		if (!group.length) return;
		this._undo.push(group);
		if (this._undo.length > this._max) this._undo.shift();
		this._redo.length = 0;
	}

	// Pops entries until one can be applied (rows that are gone are skipped).
	_step(from, to, use_old, source) {
		const sheet = this._sheet_id();
		while (from.length) {
			const group = from.pop();
			const edits = [];
			const data = this.board.list_view?.data || [];
			group.forEach((e) => {
				if (e.sheet !== sheet) return;
				const row = data.findIndex((r) => r.name === e.name);
				const col = this.board._get_col_idx(e.prop);
				if (row >= 0 && col >= 0) edits.push([row, col, use_old ? e.old_v : e.new_v]);
			});
			if (edits.length) {
				this.hot.setDataAtCell(edits, source);
				to.push(group);
				return;
			}
		}
	}
};
