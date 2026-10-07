import json
from unittest.mock import patch

import frappe
from frappe.model.document import Document
from frappe.tests.utils import FrappeTestCase

from excel_view.api import _manual_save_enabled, bulk_create_records, bulk_set_value

DOCTYPE = "EV Test Bulk Save"


def _make_doctype():
	"""A throwaway DocType with one field of each kind the server guard must tell apart."""
	if frappe.db.exists("DocType", DOCTYPE):
		frappe.delete_doc("DocType", DOCTYPE, force=True)
	frappe.get_doc(
		{
			"doctype": "DocType",
			"name": DOCTYPE,
			"module": "Excel View",
			"custom": 1,
			"autoname": "hash",
			"fields": [
				{"fieldname": "note", "fieldtype": "Data", "label": "Note"},
				# Link with a fetch_from child: the case Frappe's own link check skips.
				{"fieldname": "person", "fieldtype": "Link", "options": "User", "label": "Person"},
				{
					"fieldname": "person_name",
					"fieldtype": "Data",
					"label": "Person Name",
					"read_only": 1,
					"fetch_from": "person.full_name",
				},
				{"fieldname": "locked", "fieldtype": "Data", "label": "Locked", "read_only": 1},
				{
					"fieldname": "shown",
					"fieldtype": "Read Only",
					"label": "Shown",
					"fetch_from": "person.email",
				},
			],
			"permissions": [{"role": "System Manager", "read": 1, "write": 1, "create": 1}],
		}
	).insert(ignore_permissions=True)


def _updates(*rows):
	return json.dumps([{"name": name, "fields": fields} for name, fields in rows])


class TestBulkSetValue(FrappeTestCase):
	@classmethod
	def setUpClass(cls):
		super().setUpClass()
		frappe.set_user("Administrator")
		_make_doctype()

	@classmethod
	def tearDownClass(cls):
		frappe.set_user("Administrator")
		frappe.delete_doc("DocType", DOCTYPE, force=True)
		frappe.db.commit()
		super().tearDownClass()

	def setUp(self):
		frappe.set_user("Administrator")
		self.names = []
		for note in ("a", "b", "c"):
			doc = frappe.get_doc({"doctype": DOCTYPE, "note": note, "locked": "orig"}).insert()
			self.names.append(doc.name)
		frappe.db.commit()

	def tearDown(self):
		frappe.set_user("Administrator")
		frappe.db.delete(DOCTYPE)
		frappe.db.commit()

	def _value(self, name, fieldname):
		return frappe.db.get_value(DOCTYPE, name, fieldname)

	# ── read-only fields ──────────────────────────────────────────────────────

	def test_read_only_field_is_refused(self):
		r = bulk_set_value(DOCTYPE, _updates((self.names[0], {"locked": "changed"})))
		self.assertEqual(len(r["errors"]), 1)
		self.assertEqual(r["errors"][0]["fields"], ["locked"])
		self.assertEqual(self._value(self.names[0], "locked"), "orig")

	def test_read_only_field_with_fetch_from_is_refused(self):
		"""A fetched read-only field used to pass the guard; it must be refused."""
		r = bulk_set_value(DOCTYPE, _updates((self.names[0], {"person_name": "Made Up"})))
		self.assertEqual(len(r["errors"]), 1)
		self.assertEqual(r["errors"][0]["fields"], ["person_name"])
		self.assertFalse(self._value(self.names[0], "person_name"))

	def test_read_only_fieldtype_with_fetch_from_is_refused(self):
		"""The "Read Only" fieldtype is not flagged read_only in the meta, but is locked."""
		r = bulk_set_value(DOCTYPE, _updates((self.names[0], {"shown": "x@y.z"})))
		self.assertEqual(len(r["errors"]), 1)
		self.assertFalse(self._value(self.names[0], "shown"))

	def test_system_and_unknown_fields_are_refused(self):
		r = bulk_set_value(
			DOCTYPE,
			_updates(
				(self.names[0], {"owner": "Guest"}),
				(self.names[1], {"not_a_field": "x"}),
				(self.names[2], {"docstatus": 1}),
			),
		)
		self.assertEqual(len(r["errors"]), 3)
		self.assertEqual(self._value(self.names[0], "owner"), "Administrator")

	def test_editable_field_still_saves(self):
		r = bulk_set_value(DOCTYPE, _updates((self.names[0], {"note": "edited"})))
		self.assertEqual(r["errors"], [])
		self.assertEqual(self._value(self.names[0], "note"), "edited")

	def test_fetched_field_is_still_refreshed_by_frappe_on_a_link_edit(self):
		"""Editing the Link is allowed; Frappe itself then refreshes the fetched field."""
		full_name = frappe.db.get_value("User", "Administrator", "full_name")
		r = bulk_set_value(DOCTYPE, _updates((self.names[0], {"person": "Administrator"})))
		self.assertEqual(r["errors"], [])
		self.assertEqual(self._value(self.names[0], "person_name"), full_name)

	# ── Link values ───────────────────────────────────────────────────────────

	def test_missing_link_value_is_refused_on_update(self):
		r = bulk_set_value(DOCTYPE, _updates((self.names[0], {"person": "nobody@nowhere.invalid"})))
		self.assertEqual(len(r["errors"]), 1)
		self.assertEqual(r["errors"][0]["fields"], ["person"])
		self.assertFalse(self._value(self.names[0], "person"))

	def test_missing_link_value_is_refused_on_create(self):
		before = frappe.db.count(DOCTYPE)
		r = bulk_create_records(DOCTYPE, json.dumps([{"note": "new", "person": "nobody@nowhere.invalid"}]))
		self.assertEqual(r["created"], [])
		self.assertEqual(len(r["errors"]), 1)
		self.assertEqual(frappe.db.count(DOCTYPE), before)

	def test_valid_create_still_works(self):
		r = bulk_create_records(DOCTYPE, json.dumps([{"note": "new", "person": "Administrator"}]))
		self.assertEqual(len(r["created"]), 1)
		self.assertEqual(r["errors"], [])

	def test_old_bad_link_on_another_field_does_not_block_an_edit(self):
		"""Only the Link fields being edited are checked."""
		frappe.db.set_value(DOCTYPE, self.names[0], "person", "gone@nowhere.invalid")
		r = bulk_set_value(DOCTYPE, _updates((self.names[0], {"note": "edited"})))
		self.assertEqual(r["errors"], [])
		self.assertEqual(self._value(self.names[0], "note"), "edited")

	# ── a bad row does not stop the others ────────────────────────────────────

	def test_one_bad_row_does_not_stop_the_good_rows(self):
		a, b, c = self.names
		r = bulk_set_value(
			DOCTYPE,
			_updates(
				(a, {"note": "A2"}),
				(b, {"person": "nobody@nowhere.invalid", "note": "B2"}),
				(c, {"note": "C2"}),
			),
		)
		self.assertEqual([e["name"] for e in r["errors"]], [b])
		self.assertEqual(self._value(a, "note"), "A2")
		self.assertEqual(self._value(b, "note"), "b")
		self.assertEqual(self._value(c, "note"), "C2")

	def test_row_that_fails_after_it_was_written_is_rolled_back(self):
		"""A hook that fails after the row was written must undo that row only."""
		a, b, c = self.names
		real = Document.run_post_save_methods

		def fail_for_b(doc, *args, **kwargs):
			real(doc, *args, **kwargs)
			if doc.name == b:
				frappe.throw("hook failed after the write")

		with patch.object(Document, "run_post_save_methods", fail_for_b):
			r = bulk_set_value(
				DOCTYPE,
				_updates((a, {"note": "A2"}), (b, {"note": "B2"}), (c, {"note": "C2"})),
			)
		self.assertEqual([e["name"] for e in r["errors"]], [b])
		self.assertEqual(self._value(a, "note"), "A2")
		self.assertEqual(self._value(b, "note"), "b")
		self.assertEqual(self._value(c, "note"), "C2")

	# ── permissions ───────────────────────────────────────────────────────────

	def test_user_without_write_permission_is_refused(self):
		frappe.set_user("Guest")
		with self.assertRaises(frappe.PermissionError):
			bulk_set_value(DOCTYPE, _updates((self.names[0], {"note": "hacked"})))
		frappe.set_user("Administrator")
		self.assertEqual(self._value(self.names[0], "note"), "a")

	def test_user_without_create_permission_is_refused(self):
		frappe.set_user("Guest")
		with self.assertRaises(frappe.PermissionError):
			bulk_create_records(DOCTYPE, json.dumps([{"note": "hacked"}]))
		frappe.set_user("Administrator")
		self.assertEqual(frappe.db.count(DOCTYPE), 3)


class TestManualSaveSetting(FrappeTestCase):
	"""Manual Save changes how every user saves, so it must be off until someone opts in."""

	def setUp(self):
		self._stored = frappe.db.sql(
			"select value from tabSingles where doctype='Excel View Settings' and field='manual_save'"
		)
		frappe.db.sql("delete from tabSingles where doctype='Excel View Settings' and field='manual_save'")

	def tearDown(self):
		frappe.db.sql("delete from tabSingles where doctype='Excel View Settings' and field='manual_save'")
		if self._stored:
			frappe.db.set_single_value("Excel View Settings", "manual_save", int(self._stored[0][0] or 0))

	def test_off_when_never_configured(self):
		self.assertEqual(_manual_save_enabled(), 0)

	def test_on_only_when_an_administrator_ticks_it(self):
		frappe.db.set_single_value("Excel View Settings", "manual_save", 1)
		self.assertEqual(_manual_save_enabled(), 1)

	def test_off_again_when_unticked(self):
		frappe.db.set_single_value("Excel View Settings", "manual_save", 1)
		frappe.db.set_single_value("Excel View Settings", "manual_save", 0)
		self.assertEqual(_manual_save_enabled(), 0)

	def test_doctype_default_is_off(self):
		df = frappe.get_meta("Excel View Settings").get_field("manual_save")
		self.assertEqual(str(df.default), "0")
