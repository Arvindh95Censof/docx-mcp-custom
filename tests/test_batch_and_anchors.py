"""batch_edit, phrase-anchored comments, and minimal-diff table-cell edits."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp import server
from docx_mcp.document import DocxDocument, W
from docx_mcp.document.tracks import _flatten_para


@pytest.fixture
def doc(test_docx: Path):
    d = DocxDocument(str(test_docx))
    d.open()
    yield d
    d.close()


def _body(doc: DocxDocument) -> etree._Element:
    return doc._tree("word/document.xml")


def _accepted(doc: DocxDocument, para_id: str) -> str:
    para = doc._find_para(_body(doc), para_id)
    return "".join(s.char for s in _flatten_para(para))


def _snapshot(doc: DocxDocument) -> dict[str, bytes]:
    return {name: etree.tostring(tree) for name, tree in doc._trees.items()}


def _range_text(doc: DocxDocument, comment_id: int) -> str:
    """Visible text between a comment's range markers, in document order."""
    out: list[str] = []
    inside = False
    for el in _body(doc).iter():
        if el.tag == f"{W}commentRangeStart" and el.get(f"{W}id") == str(comment_id):
            inside = True
        elif el.tag == f"{W}commentRangeEnd" and el.get(f"{W}id") == str(comment_id):
            break
        elif inside and el.tag == f"{W}t" and el.text:
            out.append(el.text)
    return "".join(out)


def _revisions(doc: DocxDocument) -> tuple[list[str], list[str]]:
    ins = ["".join(t.text or "" for t in e.iter(f"{W}t")) for e in _body(doc).iter(f"{W}ins")]
    dels = [
        "".join(t.text or "" for t in e.iter(f"{W}delText")) for e in _body(doc).iter(f"{W}del")
    ]
    return ins, dels


# ── comment anchored to a phrase ────────────────────────────────────────────


class TestAnchoredComment:
    def test_covers_only_the_phrase(self, doc: DocxDocument):
        result = doc.add_comment("00000004", "Check this", anchor_text="30 days")
        assert result["anchor_text"] == "30 days"
        assert _range_text(doc, result["comment_id"]) == "30 days"
        assert _accepted(doc, "00000004") == "The contract term is 30 days from the effective date."

    def test_phrase_spanning_runs(self, doc: DocxDocument):
        result = doc.add_comment("00000006", "Spans runs", anchor_text="First bold")
        assert _range_text(doc, result["comment_id"]) == "First bold"
        assert _accepted(doc, "00000006") == "First bold last"

    def test_phrase_at_end_of_paragraph(self, doc: DocxDocument):
        result = doc.add_comment("00000004", "End", anchor_text="effective date.")
        assert _range_text(doc, result["comment_id"]) == "effective date."
        end = next(e for e in _body(doc).iter(f"{W}commentRangeEnd") if e.get(f"{W}id") == "0")
        assert end.getnext().find(f"{W}commentReference") is not None

    def test_without_anchor_the_whole_paragraph_is_covered(self, doc: DocxDocument):
        result = doc.add_comment("00000004", "Whole")
        assert "anchor_text" not in result
        assert _range_text(doc, result["comment_id"]) == _accepted(doc, "00000004")

    def test_not_found_changes_nothing(self, doc: DocxDocument):
        before = _snapshot(doc)
        with pytest.raises(ValueError, match="not found"):
            doc.add_comment("00000004", "x", anchor_text="no such phrase")
        assert _snapshot(doc) == before
        assert doc._tree("word/comments.xml") is None

    def test_ambiguous_phrase_needs_context(self, doc: DocxDocument):
        with pytest.raises(ValueError, match="Ambiguous"):
            doc.add_comment("00000002", "x", anchor_text="is")
        result = doc.add_comment("00000002", "x", anchor_text="is", context_before="This ")
        assert _range_text(doc, result["comment_id"]) == "is"

    def test_phrase_inside_a_tracked_insertion_widens_to_it(self, doc: DocxDocument):
        doc.replace_text("00000004", find="30", replace="45 business")
        result = doc.add_comment("00000004", "Inside ins", anchor_text="business")
        assert "business" in _range_text(doc, result["comment_id"])

    def test_phrase_ending_inside_a_tracked_insertion_widens_to_it(self, doc: DocxDocument):
        doc.replace_text("00000004", find="30", replace="45 business")
        result = doc.add_comment("00000004", "Ends mid-ins", anchor_text="45")
        assert _range_text(doc, result["comment_id"]) == "45 business"

    def test_survives_save_and_reopen(self, doc: DocxDocument, tmp_path: Path):
        doc.add_comment("00000004", "Persist", anchor_text="30 days")
        out = tmp_path / "anchored.docx"
        doc.save(str(out), backup=False)
        again = DocxDocument(str(out))
        again.open()
        try:
            assert _range_text(again, 0) == "30 days"
            assert [c["text"] for c in again.get_comments()] == ["Persist"]
        finally:
            again.close()


# ── minimal-diff table cell edit ────────────────────────────────────────────


class TestModifyCellMinimalDiff:
    def test_appended_words_are_a_pure_insertion(self, doc: DocxDocument):
        doc.modify_cell(0, 1, 0, "Row 1 A (rev)")
        assert _revisions(doc) == ([" (rev)"], [])
        assert _accepted(doc, "0000A005") == "Row 1 A (rev)"

    def test_only_the_changed_word_is_marked(self, doc: DocxDocument):
        doc.modify_cell(0, 1, 0, "Row 2 A")
        assert _revisions(doc) == (["2"], ["1"])
        assert _accepted(doc, "0000A005") == "Row 2 A"

    def test_unrelated_text_replaces_the_whole_cell(self, doc: DocxDocument):
        doc.modify_cell(0, 0, 0, "Totally different")
        assert _revisions(doc) == (["Totally different"], ["Header A"])

    def test_same_text_is_a_noop(self, doc: DocxDocument):
        before = _snapshot(doc)
        result = doc.modify_cell(0, 1, 0, "Row 1 A")
        assert result["modified"] is False
        assert _snapshot(doc) == before

    def test_empty_cell_gets_an_insertion(self, doc: DocxDocument):
        doc.modify_cell(0, 2, 1, "", tracked=False)
        doc.modify_cell(0, 2, 1, "Filled")
        assert _revisions(doc) == (["Filled"], [])
        assert _accepted(doc, "0000A009") == "Filled"

    def test_reject_restores_the_original_accept_keeps_the_edit(self, doc: DocxDocument):
        doc.modify_cell(0, 1, 0, "Row 2 A")
        doc.reject_changes()
        assert _accepted(doc, "0000A005") == "Row 1 A"
        doc.modify_cell(0, 1, 0, "Row 2 A")
        doc.accept_changes()
        assert _accepted(doc, "0000A005") == "Row 2 A"
        assert _revisions(doc) == ([], [])

    def test_editing_the_same_cell_twice_still_works(self, doc: DocxDocument):
        doc.modify_cell(0, 1, 0, "Row 1 A (first)")
        doc.modify_cell(0, 1, 0, "Row 1 A (second)")
        assert _accepted(doc, "0000A005") == "Row 1 A (second)"

    def test_untracked_edit_still_overwrites_directly(self, doc: DocxDocument):
        doc.modify_cell(0, 1, 0, "Overwritten", tracked=False)
        assert _accepted(doc, "0000A005") == "Overwritten"
        assert _revisions(doc) == ([], [])


# ── batch_edit ──────────────────────────────────────────────────────────────


class TestBatchEdit:
    def test_applies_every_operation_in_order(self, doc: DocxDocument):
        result = doc.batch_edit(
            [
                {
                    "op": "replace_text",
                    "para_id": "00000004",
                    "find": "30 days",
                    "replace": "45 days",
                },
                {
                    "op": "add_comment",
                    "para_id": "00000004",
                    "text": "Confirm",
                    "anchor_text": "45 days",
                },
                {"op": "modify_cell", "table_idx": 0, "row": 1, "col": 0, "text": "Row 1 A (rev)"},
                {"op": "insert_paragraph", "after_para_id": "00000001", "text": "Inserted"},
            ]
        )
        assert result["applied"] is True and result["count"] == 4
        assert [r["op"] for r in result["results"]] == [
            "replace_text",
            "add_comment",
            "modify_cell",
            "insert_paragraph",
        ]
        assert [r["index"] for r in result["results"]] == [0, 1, 2, 3]
        assert "45 days" in _accepted(doc, "00000004")
        assert _range_text(doc, result["results"][1]["result"]["comment_id"]) == "45 days"
        assert _accepted(doc, "0000A005") == "Row 1 A (rev)"
        assert _accepted(doc, result["results"][3]["result"]["para_id"]) == "Inserted"

    def test_batch_author_is_the_default_and_can_be_overridden(self, doc: DocxDocument):
        doc.batch_edit(
            [
                {"op": "insert_text", "para_id": "00000004", "text": " A", "position": "end"},
                {
                    "op": "insert_text",
                    "para_id": "00000002",
                    "text": " B",
                    "position": "end",
                    "author": "Someone Else",
                },
            ],
            author="Batch Author",
        )
        authors = {e.get(f"{W}author") for e in _body(doc).iter(f"{W}ins")}
        assert authors == {"Batch Author", "Someone Else"}

    def test_a_failing_operation_undoes_everything(self, doc: DocxDocument):
        before = _snapshot(doc)
        had_comments_part = (doc.workdir / "word" / "comments.xml").exists()
        with pytest.raises(ValueError, match=r"Operation 2 \(replace_text\) failed.*No changes"):
            doc.batch_edit(
                [
                    {
                        "op": "replace_text",
                        "para_id": "00000004",
                        "find": "30 days",
                        "replace": "9",
                    },
                    {"op": "add_comment", "para_id": "00000002", "text": "Creates comments part"},
                    {"op": "replace_text", "para_id": "00000004", "find": "absent", "replace": "x"},
                ]
            )
        assert _snapshot(doc) == before
        assert (doc.workdir / "word" / "comments.xml").exists() == had_comments_part
        assert doc._modified == set()

    def test_a_rolled_back_batch_leaves_a_saveable_document(
        self, doc: DocxDocument, tmp_path: Path
    ):
        with pytest.raises(ValueError):
            doc.batch_edit(
                [
                    {"op": "add_comment", "para_id": "00000002", "text": "x"},
                    {"op": "delete_paragraph", "para_id": "FFFFFFFF"},
                ]
            )
        out = tmp_path / "after_rollback.docx"
        result = doc.save(str(out), backup=False)
        assert "word/comments.xml" not in result["modified_parts"]
        import zipfile

        assert "word/comments.xml" not in zipfile.ZipFile(out).namelist()

    def test_validation_runs_before_any_edit(self, doc: DocxDocument):
        before = _snapshot(doc)
        with pytest.raises(ValueError, match="unknown op 'bogus'"):
            doc.batch_edit(
                [
                    {"op": "replace_text", "para_id": "00000004", "find": "30", "replace": "31"},
                    {"op": "bogus"},
                ]
            )
        assert _snapshot(doc) == before

    @pytest.mark.parametrize(
        ("operations", "message"),
        [
            ([], "non-empty list"),
            ("nope", "non-empty list"),
            ([["not", "an", "object"]], "must be an object"),
            ([{"op": "delete_paragraph"}], "missing para_id"),
            ([{"op": "delete_paragraph", "para_id": "x", "colour": "red"}], "unknown parameter"),
            ([{"op": "delete_paragraph", "para_id": "00000004"}] * 501, "Too many operations"),
        ],
    )
    def test_malformed_requests_are_rejected_up_front(
        self, doc: DocxDocument, operations: object, message: str
    ):
        before = _snapshot(doc)
        with pytest.raises(ValueError, match=message):
            doc.batch_edit(operations)  # type: ignore[arg-type]
        assert _snapshot(doc) == before

    def test_whole_number_floats_are_accepted_as_indexes(self, doc: DocxDocument):
        doc.batch_edit(
            [{"op": "modify_cell", "table_idx": 0.0, "row": 1.0, "col": 0.0, "text": "Row 1 A!"}]
        )
        assert _accepted(doc, "0000A005") == "Row 1 A!"

    def test_batched_edits_round_trip_through_save(self, doc: DocxDocument, tmp_path: Path):
        doc.batch_edit(
            [
                {
                    "op": "replace_text",
                    "para_id": "00000004",
                    "find": "30 days",
                    "replace": "45 days",
                },
                {
                    "op": "add_comment",
                    "para_id": "00000004",
                    "text": "Why",
                    "anchor_text": "45 days",
                },
            ]
        )
        out = tmp_path / "batched.docx"
        saved = doc.save(str(out), backup=False)
        assert saved["repairs"] == {
            "orphan_footnotes_removed": 0,
            "orphan_endnotes_removed": 0,
            "paraids_deduplicated": 0,
            "broken_rels_removed": 0,
        }

    def test_server_tool_returns_json(self, test_docx: Path):
        server.open_document(str(test_docx))
        out = json.loads(
            server.batch_edit(
                operations=[
                    {
                        "op": "replace_text",
                        "para_id": "00000004",
                        "find": "30 days",
                        "replace": "60 days",
                    }
                ],
                author="Tool",
            )
        )
        assert out["applied"] is True and out["count"] == 1
        server.close_document()
