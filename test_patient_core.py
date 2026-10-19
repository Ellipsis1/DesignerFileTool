"""Unit tests for patient_core. Run with:  python -m unittest -v"""

import json
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest import mock

import patient_core as core


def case(**overrides):
    d = dict(name_id="T. Est 1234-QWER", center="Chicago", arch_type="LAO4",
             split_file="No", stl_only="No", tooth_shade="A2",
             due_date="6/11/2026", surgery_date="", rx_date="6/3/2026",
             scan_date="6/2/2026", ios_box="IOS", designer="Byron")
    d.update(overrides)
    return d


class DateTests(unittest.TestCase):
    def test_formats(self):
        self.assertEqual(core.fmt_date("06/03/2026"), "6/03/2026")
        self.assertEqual(core.fmt_date("6/3/26", "."), "6.3.2026")
        self.assertEqual(core.fmt_date(""), "")

    def test_invalid(self):
        self.assertIsNone(core.parse_date("13/45/2026"))
        self.assertIsNone(core.parse_date("next tuesday"))

    def test_month_grid_for_date_picker(self):
        weeks = core.month_weeks(2026, 10)                # Oct 1, 2026 is a Thursday
        self.assertEqual(weeks[0], [0, 0, 0, 0, 1, 2, 3])  # Sunday first
        self.assertEqual(weeks[-1], [25, 26, 27, 28, 29, 30, 31])
        self.assertEqual(core.shift_month(2026, 12, 1), (2027, 1))
        self.assertEqual(core.shift_month(2026, 1, -1), (2025, 12))


class NamingTests(unittest.TestCase):
    def test_patient_split(self):
        p = core.patient_from("T. Est (1234-QWER)", "Milwaukee")
        self.assertEqual((p.init2, p.name3, p.uid, p.first1), ("T.", "EST", "1234-QWER", "T"))

    def test_folder_name(self):
        self.assertEqual(core.folder_name(case()), "6.3.2026 T. EST 1234-QWER Chicago")
        self.assertTrue(core.folder_name(case(arch_type="MODEL ONLY")).endswith(" MODEL ONLY"))
        self.assertEqual(core.base_folder_name(case()), "T. EST 1234-QWER Chicago")

    def test_casenotes_uses_given_today(self):
        notes = core.build_casenotes(case(), today=date(2026, 9, 29))
        self.assertIn("Awaiting Approval Date: 9/29/2026", notes)

    def test_casenotes_end_with_order_id(self):
        # below the ruled line only the Dental System order ID is left
        notes = core.build_casenotes(case(arch_type="DAO4", split_file="Double"))
        self.assertTrue(notes.endswith("Revisions: 0\n\n_______________________________________\n"
                                       "\nDental System Order ID Input\n1234-QWER_T_EST_Chicago"))
        self.assertNotIn("STL NAME", notes)


class ParseCaseNotesTests(unittest.TestCase):
    def test_round_trip(self):
        for d in (case(), case(surgery_date="6/20/2026"),
                  case(arch_type="Double - MxCD-LAO4", split_file="Double"),
                  case(arch_type="Double CD"), case(arch_type="MODEL ONLY", split_file="Upper"),
                  case(center="Atlanta - Decatur", tooth_shade="A3.5", ios_box="Box")):
            # blank fields are left out, as the form starts empty
            back = dict.fromkeys(d, "") | core.parse_casenotes(core.build_casenotes(d))
            self.assertEqual(core.build_casenotes(back), core.build_casenotes(d))
            self.assertEqual(back["arch_type"], d["arch_type"])
            self.assertEqual(back["split_file"], d["split_file"])
            self.assertEqual(core.folder_name(back), core.folder_name(d))

    def test_split_on_arch_line(self):
        # the old Excel sheet had no Split File line; the split rode on Arch Type
        text = "Unique ID: 1234-QWER\nArch Type: DAO4 Double Split File\n"
        back = core.parse_casenotes(text)
        self.assertEqual((back["arch_type"], back["split_file"]), ("DAO4", "Double"))
        self.assertEqual(core.parse_casenotes("Unique ID: 1234-QWER\nArch Type: DAO4 \n")["split_file"], "No")

    def test_fields(self):
        back = core.parse_casenotes(chr(0xFEFF) + core.build_casenotes(case()).replace("\n", "\r\n"))
        self.assertEqual(back["name_id"], "T. EST 1234-QWER")
        self.assertEqual(back["designer"], "Byron")
        self.assertEqual(back["ios_box"], "IOS")
        self.assertEqual(back["due_date"], "6/11/2026")
        self.assertNotIn("surgery_date", back)            # blank in the notes: left for the form

    def test_revisions(self):
        self.assertIn("\nRevisions: 0\n", core.build_casenotes(case()))
        notes = core.build_casenotes(case(revisions="2"))
        self.assertIn("\nRevisions: 2\n", notes)
        self.assertEqual(core.parse_casenotes(notes)["revisions"], "2")

    def test_other_text_is_rejected(self):
        self.assertEqual(core.parse_casenotes("Shopping list:\n- milk\n"), {})
        self.assertEqual(core.parse_casenotes(""), {})


def loaded_form(text):
    """The form as the window fills it from a CaseNotes text."""
    return dict.fromkeys(case(), "") | {"revisions": "0"} | core.parse_casenotes(text)


class CaseInfoChangeTests(unittest.TestCase):
    """Edits made on the form after loading CaseNotes are written into those notes."""

    def setUp(self):
        # as a designer left it: Windows line endings, a reworded line, an extra line
        self.text = (core.build_casenotes(case(revisions="1"), today=date(2026, 9, 29))
                     .replace("\n", "\r\n")
                     .replace("Designer: Byron", "Designer: Byron (covering for Zed)")
                     .replace("Revisions: 1", "Shipping: hold until Friday\r\nRevisions: 1"))
        self.form = loaded_form(self.text)

    def test_untouched_form_changes_nothing(self):
        self.assertEqual(core.casenotes_changes(self.text, self.form), [])
        self.assertEqual(core.apply_case_info(self.text, self.form), self.text)
        same_dates = self.form | {"due_date": "06/11/2026", "rx_date": "6/3/26"}   # retyped
        self.assertEqual(core.casenotes_changes(self.text, same_dates), [])
        self.assertEqual(core.casenotes_changes(self.text, self.form | {"name_id": "t. est 1234-qwer"}), [])

    def test_changed_fields_are_listed_and_written(self):
        form = self.form | {"due_date": "06/15/2026", "surgery_date": "6/20/2026", "revisions": "2"}
        self.assertEqual(core.casenotes_changes(self.text, form),
                         [("Due by Date", "6/11/2026", "6/15/2026"),
                          ("Surgery", "", "6/20/2026"),
                          ("Revisions", "1", "2")])
        want = (self.text.replace("Due by Date: 6/11/2026", "Due by Date: 6/15/2026")
                .replace("Surgery: \r\n", "Surgery: 6/20/2026\r\n")
                .replace("Revisions: 1", "Revisions: 2"))
        self.assertEqual(core.apply_case_info(self.text, form), want)   # nothing else moved
        self.assertIn("Designer: Byron (covering for Zed)\r\n", want)
        self.assertIn("Awaiting Approval Date: 9/29/2026\r\n", want)

    def test_name_arch_and_cleared_fields(self):
        form = self.form | {"name_id": "T. Est 9999-ZZZZ", "arch_type": "DAO4",
                            "split_file": "Double", "tooth_shade": ""}
        self.assertEqual(core.casenotes_changes(self.text, form),
                         [("Unique ID", "1234-QWER", "9999-ZZZZ"),
                          ("Arch Type", "LAO4", "DAO4"),
                          ("Split File", "No", "Double"),
                          ("Tooth Shade", "A2", "")])
        out = core.apply_case_info(self.text, form)
        self.assertIn("Unique ID: 9999-ZZZZ\r\n", out)
        self.assertIn("Arch Type: DAO4\r\nSplit File: Double\r\n", out)
        self.assertIn("Patient Name: T. EST\r\n", out)

    def test_missing_line_is_added_above_revisions(self):
        text = "Unique ID: 1234-QWER\nPatient Name: T. EST\nRevisions: 0\n\n____\nMODELS\n"
        out = core.apply_case_info(text, loaded_form(text) | {"due_date": "6/15/2026"})
        self.assertIn("Due by Date: 6/15/2026\nRevisions: 0\n\n____\nMODELS\n", out)


class ArchFromTreatmentTests(unittest.TestCase):
    def test_sides(self):
        self.assertEqual(core.arch_from_treatment("Lower Zirconia Arch Replacement"), "LAO4")
        self.assertEqual(core.arch_from_treatment("Upper Zirconia Arch Replacement"), "UAO4")
        self.assertEqual(core.arch_from_treatment("Upper and Lower Zirconia"), "DAO4")
        self.assertIsNone(core.arch_from_treatment(""))

    def test_kinds_per_arch(self):
        cases = {
            "Upper Partial Denture and Lower Zirconia Arch Replacement": "LAO4",
            "Upper and Lower Dentures": "Double CD",
            "Upper Denture and Lower Essential Hybrid Arch Replacement": "Double - MxCD-LAO4",
            "Upper Denture and Lower Implant Bridge": "Double - MxCD-LAO4",
            "Upper and Lower Endura Zirconia Arch Replacement": "DAO4",
            "Lower Implant Supported Denture": "Lower - MdOD",
            "Upper Overdenture and Lower Denture": "Double - MdCD-UOD",
            "Upper Zirconia Arch Replacement and Lower Overdenture": "Double - UAO4-LOD",
            "Upper Denture": "Upper - MxCD",
        }
        for pot, arch in cases.items():
            with self.subTest(pot):
                self.assertEqual(core.arch_from_treatment(pot), arch)

    def test_partials_are_left_blank(self):
        self.assertIsNone(core.arch_from_treatment("Upper and Lower Partial Dentures"))
        self.assertIsNone(core.arch_from_treatment("Upper Partial Denture"))

    def test_arch_type_fields(self):
        self.assertEqual(core.arch_from_treatment("Zirconia Arch Replacement", "Monoblock",
                                                  "Monoblock"), "DAO4")
        # a filled-in field doesn't bring back an arch the plan calls partial
        self.assertEqual(core.arch_from_treatment(
            "Upper Partial Denture and Lower Zirconia Arch Replacement", "Monoblock",
            "Monoblock"), "LAO4")

    def test_renamed_type_falls_back(self):
        saved = list(core.ARCH_TYPES)
        try:
            core.ARCH_TYPES.remove("Double CD")
            self.assertEqual(core.arch_from_treatment("Upper and Lower Dentures"), "DAO4")
        finally:
            core.ARCH_TYPES[:] = saved


class GuessCenterTests(unittest.TestCase):
    CENTERS = ["Charlotte", "Charlotte - University City", "Columbia MD", "Columbia SC (Metro)",
               "Columbus", "St Louis", "Woburn"]

    def guess(self, text):
        with mock.patch.object(core, "CENTERS", self.CENTERS):
            return core.guess_center(text)

    def test_state_tells_similar_names_apart(self):
        self.assertEqual(self.guess("Dental Implants 1 Main St Columbia, South Carolina 29201\n"
                                    "Lab Rx\n"), "Columbia SC (Metro)")
        self.assertEqual(self.guess("Practice 2 Elm Rd Columbia, MD 21044\nLab Rx\n"),
                         "Columbia MD")

    def test_address_line_comes_first(self):
        text = ("Implant Dentistry Associates of Metro Charlotte, P.C. Fort Mill, South Carolina\n"
                "Lab Rx\nNote: same as the Woburn case\n")
        self.assertEqual(self.guess(text), "Charlotte")
        self.assertEqual(self.guess("Practice, Woburn, Massachusetts\nLab Rx\n"), "Woburn")

    def test_punctuation_and_misses(self):
        self.assertEqual(self.guess("Clinic, St. Louis, Missouri\nLab Rx\n"), "St Louis")
        self.assertIsNone(self.guess("Clinic, Boston, Massachusetts\nLab Rx\n"))


class FakePage:
    """Stands in for a pdfplumber page: words are (text, x0, top)."""

    def __init__(self, *words):
        self.words = [dict(text=t, x0=x, top=top) for t, x, top in words]

    def extract_words(self):
        return self.words

    def extract_text(self):
        return " ".join(w["text"] for w in self.words)


class RxPdfTests(unittest.TestCase):
    # Layout of the Lab Rx: labels at x=55, values from x~130, the Note column from x=355
    # and a couple of points higher than the form rows beside it.
    PAGE1 = FakePage(
        ("Patient", 55, 84.4), ("T.", 115, 82.0), ("Est", 131, 82.0), ("1234-QWER", 150, 82.0),
        ("Due", 362, 84.4), ("Date:", 383, 84.4), ("6/8/2026", 490, 84.4),
        ("Plan", 55, 148.1), ("of", 78, 148.1), ("Upper", 151, 148.1), ("Zirconia", 181, 148.1),
        ("Arch", 218, 148.1), ("shade", 355, 145.9), ("lower", 400, 145.9),
        ("Treatment:", 55, 159.4), ("Replacement", 151, 159.4), ("BL1", 355, 157.1),
        ("Surgical", 55, 367.9), ("Arch:", 93, 367.9), ("STL", 143, 367.9), ("Only", 163, 367.9),
    )
    PAGE2 = FakePage(
        ("Tooth", 55, 216.4), ("Shade:", 82, 216.4), ("BL3", 140, 216.4),
        ("Tooth", 55, 242.6), ("Shade", 82, 242.6), ("Other:", 114, 242.6),
        ("Scan", 55, 470.6), ("Type:", 80, 470.6), ("Intraoral", 131, 470.6),
    )

    def parse(self, *pages):
        pdf = mock.MagicMock()
        pdf.__enter__.return_value.pages = list(pages)
        with mock.patch.object(core, "PDF_AVAILABLE", True), \
                mock.patch.object(core, "pdfplumber", create=True) as plumber:
            plumber.open.return_value = pdf
            return core.parse_rx_pdf("rx.pdf")

    def test_rows_read_left_to_right(self):
        rows = [" ".join(w["text"] for w in r) for r in core._rows(self.PAGE1)]
        self.assertIn("Plan of Upper Zirconia Arch shade lower", rows)

    def test_note_column_does_not_hide_the_form(self):
        out = self.parse(self.PAGE1, self.PAGE2)
        self.assertEqual(out["name_id"], "T. Est 1234-QWER")
        self.assertEqual(out["due_date"], "6/8/2026")
        self.assertEqual(out["plan_of_treatment"], "Upper Zirconia Arch Replacement")
        self.assertEqual(out["arch_type"], "UAO4")      # "lower" in the note is not the plan
        self.assertEqual(out["stl_only"], "Yes")

    def test_fields_pushed_onto_page_two_are_read(self):
        out = self.parse(self.PAGE1, self.PAGE2)
        self.assertEqual(out["tooth_shade"], "BL3")
        self.assertEqual(out["ios_box"], "IOS")


class GuessRoleTests(unittest.TestCase):
    def test_by_extension(self):
        self.assertEqual(core.guess_role("shot 1.PNG"), "SCREENSHOT")
        self.assertEqual(core.guess_role("export.ZIP"), "KEEP")
        self.assertEqual(core.guess_role("upper bite.zip", "UAO4"), "KEEP")   # name tokens don't matter

    def test_by_name(self):
        self.assertEqual(core.guess_role("upper_jaw.stl"), "MX")
        self.assertEqual(core.guess_role("Mandibular.stl"), "MD")
        self.assertEqual(core.guess_role("bite scan.stl"), "BITE")
        self.assertEqual(core.guess_role("lower cutback.stl"), "MD_CUTBACK")
        self.assertEqual(core.guess_role("upper base.stl"), "MX_BASE")
        self.assertEqual(core.guess_role("lower teeth.stl"), "MD_TEETH")
        self.assertEqual(core.guess_role("upper model.stl"), "MDL_MX")
        self.assertEqual(core.guess_role("lower working.stl"), "WRK_MD")
        self.assertEqual(core.guess_role("upper.dcm"), "MX_DCM")
        self.assertEqual(core.guess_role("upper working.dcm"), "WRK_MX")

    def test_falls_back_on_arch(self):
        self.assertEqual(core.guess_role("scan.stl", "LAO4"), "MD")
        self.assertEqual(core.guess_role("scan.stl", "UAO4"), "MX")
        self.assertIsNone(core.guess_role("scan.stl", "DAO4"))   # both arches: ask
        self.assertIsNone(core.guess_role("scan.stl"))

    def test_md_inside_word_is_not_lower(self):
        # "cmd" contains "md" but is not a lower-arch token
        self.assertIsNone(core.guess_role("cmd_export.stl"))


class FileStageTests(unittest.TestCase):
    BASE = Path("Desktop") / "T. REV 1234-QWER Chicago"

    def stage(self, rel, base=BASE):
        return core.file_stage(self.BASE / rel, base)

    def test_from_folder(self):
        for folder, stage in (("Revision 1", 1), ("Rev 2", 2), ("rev2", 2), ("Redesign 3", 3),
                              ("Re-design_4", 4), ("Revision", 1), ("Revised files", 1),
                              ("Revision 2 screenshots", 2), ("Initial Design", 0),
                              ("Original", 0), ("Screenshots", 0), ("Review", 0)):
            self.assertEqual(self.stage(f"{folder}/upper.dcm"), stage, folder)
        self.assertEqual(self.stage("upper.dcm"), 0)
        self.assertEqual(self.stage("Rev 2/Screenshots/a.png"), 2)
        self.assertEqual(self.stage("Initial Design/upper rev 1.dcm"), 0)    # the folder decides

    def test_from_file_name(self):
        for name, stage in (("upper Rev 1.dcm", 1), ("upper_rev2.dcm", 2), ("Revision 3 lower.png", 3),
                            ("upper redesign-2.dcm", 2), ("upper.dcm", 0), ("preview 1.png", 0),
                            ("1234-QWER_MX_T_REV.dcm", 0), ("upper revised.dcm", 0)):   # needs a number
            self.assertEqual(self.stage(name), stage, name)

    def test_working_folder_name_is_not_read(self):
        # 'T. REV 1234-...' is the patient, not a revision, with or without a known base.
        self.assertEqual(self.stage("upper.dcm"), 0)
        self.assertEqual(self.stage("upper.dcm", base=None), 0)
        self.assertEqual(self.stage("Rev 1/upper.dcm", base=None), 1)
        self.assertEqual(self.stage("Rev 1/upper.dcm", base=Path("elsewhere")), 1)

    def test_stage_name(self):
        self.assertEqual([core.stage_name(n) for n in (0, 1, 12)],
                         ["Initial Design", "Revision 1", "Revision 12"])


class PlanAndCopyTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.src = self.root / "src"
        self.src.mkdir()
        self.main = self.root / "Desktop" / core.folder_name(case())

    def tearDown(self):
        self.tmp.cleanup()

    def make(self, name, text="x"):
        p = self.src / name
        p.write_text(text)
        return p

    def test_names_and_routing(self):
        items = [(self.make("a.stl"), "MD"), (self.make("b.png"), "SCREENSHOT"),
                 (self.make("m.stl"), "MDL_MX"), (self.make("notes.pdf"), "KEEP"),
                 (self.make("d.stl"), None)]
        plan = core.plan_file_copies(items, case(), self.main)
        rel = [[p.relative_to(self.main).as_posix() for p in dests] if dests else None
               for dests in plan]
        self.assertEqual(rel, [
            ["1234-QWER STL/1234-QWER_MD_T_EST_Chicago.stl"],
            ["Design Screenshots/b.png"],                              # screenshots keep their name
            ["1234-QWER STL/1234-QWER_MDL_MX_T_EST_Chicago.stl",       # models go to both
             "3D Viewer/1234-QWER_MDL_MX_T_EST_Chicago.stl"],
            ["notes.pdf"],
            None,
        ])

    def test_non_stl_for_stl_folder(self):
        items = [(Path("a.STL"), "MD"), (Path("b.ply"), "MX"), (Path("m.dcm"), "MDL_MX"),
                 (Path("w.dcm"), "WRK_MX"), (Path("s.png"), "SCREENSHOT"), (Path("d.obj"), None)]
        self.assertEqual(core.non_stl_for_stl_folder(items), [Path("b.ply"), Path("m.dcm")])

    def test_zips_not_for_patient(self):
        ok = ["1234-QWER export.zip", "scan_1234qwer.ZIP", "1234_QWER.zip",     # unique ID
              "1234.zip", "scan1234_v2.zip", "qwer export.zip", "Export_QWER.zip",  # half the ID
              "Estrada_Tom.zip", "TomEstrada.zip", "TEstrada.zip", "T EST scans.zip"]  # name
        stray = ["export.zip", "5678-ASDF.zip", "latest.zip", "Smith_John.zip",
                 "20261234.zip", "qwerty.zip", "123-QWE.zip"]
        items = [(Path(n), "KEEP") for n in ok + stray]
        items += [(Path("export.stl"), "MD"), (Path("unassigned.zip"), None)]
        self.assertEqual(core.zips_not_for_patient(items, case()), [Path(n) for n in stray])

    def test_zip_keeps_its_name(self):
        src = self.make("1234-QWER Export.zip")
        plan = core.plan_file_copies([(src, core.guess_role(src))], case(), self.main)
        self.assertEqual(plan, [[self.main / "1234-QWER Export.zip"]])

    def test_collisions_in_batch_and_on_disk(self):
        existing = self.main / "1234-QWER STL" / "1234-QWER_MD_T_EST_Chicago.stl"
        existing.parent.mkdir(parents=True)
        existing.write_text("old")
        shots = self.main / "Design Screenshots"
        shots.mkdir()
        (shots / "s1.png").write_text("old")

        items = [(self.make("a.stl"), "MD"), (self.make("b.stl"), "MD"),
                 (self.make("s1.png"), "SCREENSHOT"), (self.make("s2.png"), "SCREENSHOT")]
        names = [dests[0].name for dests in core.plan_file_copies(items, case(), self.main)]
        self.assertEqual(names, [
            "1234-QWER_MD_T_EST_Chicago_02.stl",
            "1234-QWER_MD_T_EST_Chicago_03.stl",
            "s1_02.png",
            "s2.png",
        ])

    def test_running_again_skips_files_already_there(self):
        dcm, model, shot = self.make("upper.dcm"), self.make("upper model.stl"), self.make("s1.png")
        items = [(dcm, "MX_DCM"), (model, "MDL_MX"), (shot, "SCREENSHOT")]
        core.copy_files(zip([dcm, model, shot], core.plan_file_copies(items, case(), self.main)))
        self.assertEqual(core.plan_file_copies(items, case(), self.main), [[], [], []])

        dcm.write_text("a changed design")                 # a changed file is still never overwritten
        (self.main / "3D Viewer" / "1234-QWER_MDL_MX_T_EST_Chicago.stl").unlink()
        plan = core.plan_file_copies(items, case(), self.main)
        self.assertEqual(plan, [[self.main / "3D Viewer" / "1234-QWER_MX_T_EST_02.dcm"],
                                [self.main / "3D Viewer" / "1234-QWER_MDL_MX_T_EST_Chicago.stl"],
                                []])

    def rel(self, paths):
        return [p.relative_to(self.main).as_posix() for p in paths]

    def run_round(self, items):
        """Create for one round, as the GUI does: moves first, then copies."""
        stages = [core.file_stage(src, self.src) for src, _ in items]
        plan = core.plan_file_copies(items, case(), self.main, stages)
        moves = core.plan_folder_moves(items, case(), self.main, stages)
        core.move_files(moves)
        core.copy_files(zip([src for src, _ in items], plan))
        return [self.rel(dests) for dests in plan], [(self.rel([a])[0], self.rel([b])[0]) for a, b in moves]

    def tree(self, sub):
        return sorted(p.relative_to(self.main / sub).as_posix()
                      for p in (self.main / sub).rglob("*") if p.is_file())

    def test_revisions(self):
        for n, text in ((1, "rev 1"), (2, "rev 2")):
            (self.src / f"Revision {n}").mkdir()
            (self.src / f"Revision {n}" / "upper.dcm").write_text(text)
        (self.src / "Revision 1" / "s2.png").write_text("shot 2")
        first = [(self.make("upper.dcm"), "MX_DCM"), (self.make("lower.dcm"), "MD_DCM"),
                 (self.make("s1.png"), "SCREENSHOT"), (self.make("upper model.stl"), "MDL_MX"),
                 (self.make("1234 export.zip"), "KEEP")]
        rev1 = first + [(self.src / "Revision 1" / "upper.dcm", "MX_DCM"),
                        (self.src / "Revision 1" / "s2.png", "SCREENSHOT")]
        rev2 = rev1 + [(self.src / "Revision 2" / "upper.dcm", "MX_DCM")]

        # The first send, before any revision, is flat.
        plan, moves = self.run_round(first)
        self.assertEqual(plan[:3], [["3D Viewer/1234-QWER_MX_T_EST.dcm"],
                                    ["3D Viewer/1234-QWER_MD_T_EST.dcm"],
                                    ["Design Screenshots/s1.png"]])
        self.assertEqual(moves, [])

        # Revision 1 replaces the upper: the initial one moves down, the new one goes on top
        # as _REV_1. Screenshots split by round. Nothing already there is copied again.
        plan, moves = self.run_round(rev1)
        self.assertEqual(moves, [
            ("Design Screenshots/s1.png", "Design Screenshots/Initial Design/s1.png"),
            ("3D Viewer/1234-QWER_MX_T_EST.dcm", "3D Viewer/Initial Design/1234-QWER_MX_T_EST.dcm"),
        ])
        self.assertEqual(plan, [[], [], [], [], [],
                                ["3D Viewer/1234-QWER_MX_T_EST_REV_1.dcm"],
                                ["Design Screenshots/Revision 1/s2.png"]])

        # Revision 2 replaces the upper again: REV_1 moves into Revision 1.
        plan, moves = self.run_round(rev2)
        self.assertEqual(moves, [("3D Viewer/1234-QWER_MX_T_EST_REV_1.dcm",
                                  "3D Viewer/Revision 1/1234-QWER_MX_T_EST_REV_1.dcm")])
        self.assertEqual(plan, [[]] * 7 + [["3D Viewer/1234-QWER_MX_T_EST_REV_2.dcm"]])
        self.assertEqual(self.tree("3D Viewer"), [
            "1234-QWER_MDL_MX_T_EST_Chicago.stl",
            "1234-QWER_MD_T_EST.dcm",                       # never replaced: stays on top
            "1234-QWER_MX_T_EST_REV_2.dcm",
            "Initial Design/1234-QWER_MX_T_EST.dcm",
            "Revision 1/1234-QWER_MX_T_EST_REV_1.dcm",
        ])
        self.assertEqual(self.tree("Design Screenshots"), ["Initial Design/s1.png",
                                                           "Revision 1/s2.png"])
        self.assertEqual((self.main / "3D Viewer" / "1234-QWER_MX_T_EST_REV_2.dcm").read_text(),
                         "rev 2")
        self.assertTrue(all(src.exists() for src, _ in rev2))     # the working folder is untouched

    def test_moves_never_land_on_a_file(self):
        shots = self.main / "Design Screenshots"
        (shots / "Initial Design").mkdir(parents=True)
        (shots / "s1.png").write_text("first send")
        (shots / "Initial Design" / "s1.png").write_text("put there by hand")
        (self.src / "Revision 1").mkdir()
        (self.src / "Revision 1" / "s2.png").write_text("shot 2")
        plan, moves = self.run_round([(self.src / "Revision 1" / "s2.png", "SCREENSHOT")])
        self.assertEqual(moves, [("Design Screenshots/s1.png",
                                  "Design Screenshots/Initial Design/s1_02.png")])
        self.assertEqual(self.tree("Design Screenshots"), [
            "Initial Design/s1.png", "Initial Design/s1_02.png", "Revision 1/s2.png"])

    def test_copy_keeps_originals(self):
        src = self.make("upper.stl", "mesh")
        core.create_folder_structure(self.main, "1234-QWER", notes="hello")
        plan = core.plan_file_copies([(src, "MX")], case(), self.main)
        core.copy_files(zip([src], plan))
        self.assertTrue(src.exists())
        self.assertEqual(plan[0][0].read_text(), "mesh")
        self.assertEqual((self.main / "CaseNotes.txt").read_text(), "hello")
        for sub in ("3D Viewer", "Design Screenshots", "1234-QWER STL"):
            self.assertTrue((self.main / sub).is_dir())

    def test_rx_pdf_goes_beside_casenotes(self):
        rx = self.make("Lab Rx.PDF", "rx")
        self.assertEqual(core.casenotes_name(case()), "1234-QWER_T_EST_CaseNotes.txt")
        core.create_folder_structure(self.main, "1234-QWER", notes="hello",
                                     notes_name=core.casenotes_name(case()))
        self.assertEqual((self.main / "1234-QWER_T_EST_CaseNotes.txt").read_text(), "hello")
        plan = core.plan_rx_copy(rx, self.main, case())
        self.assertEqual(plan, [self.main / "1234-QWER_T_EST_Rx.pdf"])
        core.copy_files([(rx, plan)])
        self.assertTrue(rx.exists())                              # copied, not moved
        self.assertEqual(core.plan_rx_copy(rx, self.main, case()), [])   # same file again: nothing to do
        rx.write_text("revised rx")
        self.assertEqual(core.plan_rx_copy(rx, self.main, case()),
                         [self.main / "1234-QWER_T_EST_Rx_02.pdf"])

    def test_final_casenotes_are_the_case_information_only(self):
        full = core.build_casenotes(case(), today=date(2026, 9, 29))
        head = core.casenotes_header(full)
        self.assertTrue(head.startswith("Center: Chicago\nPatient Name: T. EST\n"))
        self.assertTrue(head.endswith("Awaiting Approval Date: 9/29/2026\nRevisions: 0\n"))
        self.assertEqual(core.casenotes_header(head), head)       # already trimmed: unchanged
        self.assertEqual(core.parse_casenotes(head), core.parse_casenotes(full))

    def test_final_casenotes_from_a_loaded_file(self):
        name = core.casenotes_name(case())
        # an old Excel-era file: cp1252, Windows line endings, edited by hand
        head = "Unique ID: 1234-QWER\r\nDesigner: Zoë (was Byron)\r\nDue by Date: 6/11/2026\r\n"
        raw = (head + "Revisions: 1\r\n\r\n____________\r\nMAX STL NAME\r\ncafé note\r\n").encode("cp1252")
        working = self.src / "CaseNotes.txt"
        working.write_bytes(raw)
        text, enc = core.read_casenotes(working)
        self.assertEqual((enc, core.parse_casenotes(text)["designer"]), ("cp1252", "Zoë (was Byron)"))
        form = loaded_form(text)
        self.assertEqual(core.casenotes_changes(text, form), [])
        form |= {"revisions": "2", "due_date": "6/15/2026"}
        dest = core.write_final_casenotes(working, self.main, name, form)
        self.assertEqual(dest, self.main / name)
        new_head = head.replace("6/11/2026", "6/15/2026") + "Revisions: 2\r\n"
        self.assertEqual(dest.read_bytes(), new_head.encode("cp1252"))
        # the working copy stays, complete, with the same changes
        self.assertEqual(working.read_bytes(), raw.replace(b"Revisions: 1", b"Revisions: 2")
                                                  .replace(b"6/11/2026", b"6/15/2026"))
        # dropping the finished folder's own notes back in just updates them
        core.write_final_casenotes(dest, self.main, name, form | {"revisions": "3"})
        self.assertEqual(dest.read_bytes(), new_head.replace("Revisions: 2", "Revisions: 3").encode("cp1252"))

    def test_initial_folder_gets_notes_and_the_rx_moved_in(self):
        base = self.root / "Desktop" / core.base_folder_name(case())
        rx = self.make("Lab Rx.pdf", "rx")
        notes, moved = core.create_initial_folder(base, case(), rx)
        # the "!" keeps the notes above anything the designer saves in the folder
        self.assertEqual(notes, base / "!1234-QWER_T_EST_CaseNotes.txt")
        self.assertEqual(notes.read_text(), core.build_casenotes(case()))
        self.assertEqual(moved, base / "1234-QWER_T_EST_Rx.pdf")
        self.assertEqual(moved.read_text(), "rx")
        self.assertFalse(rx.exists())                             # moved, not copied
        self.assertEqual(sorted(p.name for p in base.iterdir()), [notes.name, moved.name])
        # the finished folder's copies have no "!"
        self.assertEqual(core.plan_rx_copy(moved, self.main, case()),
                         [self.main / "1234-QWER_T_EST_Rx.pdf"])
        self.assertEqual(core.write_final_casenotes(notes, self.main, core.casenotes_name(case()),
                                                    loaded_form(notes.read_text())),
                         self.main / "1234-QWER_T_EST_CaseNotes.txt")
        self.assertEqual(core.create_initial_folder(base, case(), moved), (notes, None))   # again
        self.assertTrue(moved.exists())
        self.assertEqual(core.create_initial_folder(base, case()), (notes, None))          # no Rx

    def test_working_folder_files(self):
        base = self.root / "Desktop" / core.base_folder_name(case())
        notes, rx = core.create_initial_folder(base, case(), self.make("Lab Rx.pdf", "rx"))
        for rel in ("upper.stl", "shots/front.png", "desktop.ini", "~lock.tmp",
                    "3D Viewer/1234-QWER_MX_T_EST.dcm", "1234-QWER STL/old.stl", "other.pdf"):
            (base / rel).parent.mkdir(exist_ok=True)
            (base / rel).write_text("x")
        found_rx, files = core.working_folder_files(notes, case())
        self.assertEqual(found_rx, rx)
        self.assertEqual([f.relative_to(base).as_posix() for f in files],
                         ["other.pdf", "shots/front.png", "upper.stl"])
        # CaseNotes sitting somewhere not named for the patient: leave the neighbours alone
        loose = self.make("1234-QWER_T_EST_CaseNotes.txt")
        self.make("someone elses.stl")
        self.assertEqual(core.working_folder_files(loose, case()), (None, []))

    def test_safe_name_strips_windows_chars(self):
        self.assertEqual(core.safe_name('a<b>:c"d?e*'), "abcde")


class SettingsTests(unittest.TestCase):
    def test_round_trip_and_bad_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "sub" / "settings.json"
            self.assertEqual(core.load_settings(path), {})          # missing file
            core.save_settings({"designer": "Josh"}, path)
            self.assertEqual(core.load_settings(path), {"designer": "Josh"})
            path.write_text("{not json")
            self.assertEqual(core.load_settings(path), {})          # corrupt file


class ListsTests(unittest.TestCase):
    """apply_lists mutates module lists, so snapshot and restore them around each test."""

    def setUp(self):
        self.saved = {k: list(v) for k, v in core._EDITABLE.items()}
        self.saved_arch_types = list(core.ARCH_TYPES)
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / core.LISTS_FILE

    def tearDown(self):
        for k, v in self.saved.items():
            core._EDITABLE[k][:] = v
        core.ARCH_TYPES[:] = self.saved_arch_types
        self.tmp.cleanup()

    def write(self, text):
        self.path.write_text(text, encoding="utf-8")
        return self.path

    def test_apply_updates_in_place_and_ignores_bad_keys(self):
        ref = core.DESIGNERS
        p = self.write(json.dumps({"version": "v2", "designers": ["Zed"], "centers": "nope"}))
        self.assertEqual(core.apply_lists([p]), (p, "v2"))
        self.assertIs(core.DESIGNERS, ref)              # same object, new contents
        self.assertEqual(core.DESIGNERS, ["Zed"])
        self.assertEqual(core.CENTERS, self.saved["centers"])  # invalid value ignored

    def test_arch_groups_come_from_json(self):
        ref = core.ARCH_TYPES
        shades = list(core.TOOTH_SHADES)
        p = self.write(json.dumps({"upper_arch": ["New Upper"], "double_arch": ["New Double"],
                                   "other_arch": ["New Other"], "tooth_shades": ["Z9"]}))
        core.apply_lists([p])
        self.assertIs(core.ARCH_TYPES, ref)             # the dropdown's list, rebuilt in place
        self.assertEqual(core.ARCH_TYPES,
                         ["New Upper"] + self.saved["lower_arch"] + ["New Double", "New Other"])
        self.assertEqual((core.is_upper("New Upper"), core.is_lower("New Upper")), (True, False))
        self.assertEqual((core.is_upper("New Double"), core.is_lower("New Double")), (True, True))
        self.assertEqual((core.is_upper("New Other"), core.is_lower("New Other")), (True, True))
        self.assertFalse(core.is_upper("UAO4"))         # no longer offered
        self.assertEqual(core.guess_role("scan.stl", "New Upper"), "MX")
        self.assertEqual(core.TOOTH_SHADES, shades)     # not editable from the JSON

    def test_dropdown_values_must_be_in_their_lists(self):
        p = self.write(json.dumps({"designers": ["Byron"], "centers": ["Chicago", "Atlanta - Decatur"]}))
        core.apply_lists([p])
        self.assertEqual(core.match_choice(" chicago ", core.CENTERS), "Chicago")
        self.assertIsNone(core.match_choice("Chicgo", core.CENTERS))
        self.assertEqual(core.invalid_choices(case(center="chicago", designer="BYRON")), [])
        self.assertEqual(core.invalid_choices(case(tooth_shade="", ios_box="")), [])   # blank is the caller's call
        self.assertEqual(core.invalid_choices(case(center="Chicgo", arch_type="LAO5", split_file="Yes",
                                                   tooth_shade="Z9", stl_only="Maybe", ios_box="Mail")),
                         [("center", "Center", "Chicgo"), ("arch_type", "Arch Type", "LAO5"),
                          ("split_file", "Split File", "Yes"), ("tooth_shade", "Tooth Shade", "Z9"),
                          ("stl_only", "STL Only", "Maybe"), ("ios_box", "IOS or Box", "Mail")])
        self.assertEqual(core.filter_choices("decat", core.CENTERS), ["Atlanta - Decatur"])
        self.assertEqual(core.filter_choices(" A", core.CENTERS), core.CENTERS)       # in both
        self.assertEqual(core.filter_choices("chicago", core.CENTERS), core.CENTERS)  # picked: full list
        self.assertEqual(core.filter_choices("", core.CENTERS), core.CENTERS)
        self.assertEqual(core.filter_choices("zz", core.CENTERS), [])
        self.assertEqual(core.filter_choices("mx", ["Upper - MxOD", "LAO4", "Double - MxCD-LAO4"]),
                         ["Upper - MxOD", "Double - MxCD-LAO4"])
        core.apply_lists([self.write(json.dumps({"designers": ["Zed"]}))])   # File > Reload Lists
        self.assertEqual(core.invalid_choices(case()), [("designer", "Designer", "Byron")])

    def test_load_uses_file_beside_app(self):
        p = self.write(json.dumps({"version": "b1", "designers": ["Bo"]}))
        with mock.patch.object(core, "default_lists_path", return_value=p):
            self.assertEqual(core.load_lists(), (p, "b1"))
        self.assertEqual(core.DESIGNERS, ["Bo"])

    def test_missing_or_broken_file_uses_built_in(self):
        with mock.patch.object(core, "default_lists_path", return_value=self.path):
            self.assertEqual(core.load_lists(), (None, None))    # no file
            self.write("{broken json")
            self.assertEqual(core.load_lists(), (None, None))
        self.assertEqual(core.DESIGNERS, self.saved["designers"])


if __name__ == "__main__":
    unittest.main()
