"""
Integration tests for preppy/pco.py — PCO API routes.
External PCO API calls are mocked via the `responses` library.
Requires a running Postgres (DATABASE_URL).
"""

import json
import pytest
import responses
from tests.conftest import requires_db

import re as _re

PCO_API = "https://api.planningcenteronline.com"
PCO_UPLOAD = "https://upload.planningcenteronline.com/v2/files"

# Pattern to catch any arrangements/.../attachments request (used by _sections_from_pco_pdf)
_ATTACHMENTS_RE = _re.compile(
    r"https://api\.planningcenteronline\.com/services/v2/songs/.+/arrangements/.+/attachments"
)


def _mock_all_attachments():
    """Mock the attachments endpoint for any song/arrangement combo (returns empty)."""
    responses.get(_ATTACHMENTS_RE, json={"data": []})


def _mock_service_types():
    responses.get(
        f"{PCO_API}/services/v2/service_types",
        json={"data": [
            {"id": "111", "attributes": {"name": "Sunday AM"}},
        ]},
    )


def _mock_plans():
    responses.get(
        f"{PCO_API}/services/v2/service_types/111/plans",
        json={"data": [
            {
                "id": "plan-1",
                "attributes": {
                    "title": "March 8 Service",
                    "series_title": "",
                    "sort_date": "2026-03-08T09:00:00Z",
                    "items_count": 5,
                },
            },
        ]},
    )


def _mock_plan_items(service_type_id="111", plan_id="plan-1"):
    responses.get(
        f"{PCO_API}/services/v2/service_types/{service_type_id}/plans/{plan_id}/items",
        json={
            "data": [
                {
                    "id": "item-1",
                    "attributes": {"item_type": "header", "title": "Opening", "sequence": 0},
                    "relationships": {},
                },
                {
                    "id": "item-2",
                    "attributes": {"item_type": "song", "title": "Amazing Grace", "sequence": 1},
                    "relationships": {
                        "song": {"data": {"id": "pco-song-1"}},
                        "arrangement": {"data": {"id": "pco-arr-1"}},
                    },
                },
                {
                    "id": "item-3",
                    "attributes": {"item_type": "song", "title": "How Great", "sequence": 2},
                    "relationships": {
                        "song": {"data": {"id": "pco-song-2"}},
                        "arrangement": {"data": {"id": "pco-arr-2"}},
                    },
                },
            ],
            "included": [
                {"id": "pco-song-1", "type": "Song", "attributes": {"title": "Amazing Grace", "author": "John Newton"}},
                {"id": "pco-song-2", "type": "Song", "attributes": {"title": "How Great Thou Art", "author": "Stuart Hine"}},
                {"id": "pco-arr-1", "type": "Arrangement", "attributes": {"name": "Main", "chord_chart_key": "G", "bpm": 72}},
                {"id": "pco-arr-2", "type": "Arrangement", "attributes": {"name": "Main", "chord_chart_key": "Bb", "bpm": 68}},
            ],
        },
    )


def _mock_arrangement_detail(song_id, arr_id, sequence=None, chord_chart="", attachments=None):
    responses.get(
        f"{PCO_API}/services/v2/songs/{song_id}/arrangements/{arr_id}",
        json={"data": {"attributes": {
            "sequence": sequence or [],
            "chord_chart": chord_chart or "",
        }}},
    )
    # Mock the attachments endpoint (used by _sections_from_pco_pdf)
    responses.get(
        f"{PCO_API}/services/v2/songs/{song_id}/arrangements/{arr_id}/attachments",
        json={"data": attachments or []},
    )


@requires_db
class TestListPlans:
    @responses.activate
    def test_list_plans(self, db_app):
        client, uid = db_app
        _mock_service_types()
        _mock_plans()

        resp = client.get("/api/pco/plans")
        assert resp.status_code == 200
        plans = resp.get_json()
        assert len(plans) == 1
        assert plans[0]["id"] == "plan-1"
        assert plans[0]["serviceTypeName"] == "Sunday AM"
        assert plans[0]["date"] == "2026-03-08"


@requires_db
class TestImportPlan:
    @responses.activate
    def test_import_creates_setlist_with_headers(self, db_app):
        client, uid = db_app
        _mock_plan_items()
        _mock_arrangement_detail("pco-song-1", "pco-arr-1", ["Intro", "V1", "C1"])
        _mock_arrangement_detail("pco-song-2", "pco-arr-2", ["V1", "C1", "Bridge"])

        resp = client.post("/api/pco/import/plan-1", json={
            "serviceTypeId": "111",
            "date": "2026-03-08",
            "title": "March 8 Service",
        })
        assert resp.status_code == 201
        data = resp.get_json()
        assert "setlistId" in data

        setlists = client.get("/api/setlists").get_json()
        assert len(setlists) >= 1
        sl = next(s for s in setlists if s["pco_plan_id"] == "plan-1")
        assert sl["pco_service_type_id"] == "111"

        items = sl["items"]
        assert items[0]["itemType"] == "header"
        assert items[0]["label"] == "Opening"
        assert items[1]["itemType"] == "song"
        assert items[1]["title"] == "Amazing Grace"
        assert items[2]["itemType"] == "song"
        assert items[2]["title"] == "How Great Thou Art"

    @responses.activate
    def test_import_populates_sections_from_sequence(self, db_app):
        """Sequence labels like 'Verse 1' are converted to shorthand 'V1'."""
        client, uid = db_app
        _mock_plan_items()
        _mock_arrangement_detail("pco-song-1", "pco-arr-1", ["Intro", "Verse 1", "Chorus 1"])
        _mock_arrangement_detail("pco-song-2", "pco-arr-2")

        client.post("/api/pco/import/plan-1", json={
            "serviceTypeId": "111", "date": "2026-03-08", "title": "Test",
        })

        songs = client.get("/api/songs").get_json()
        grace = next(s for s in songs if s["title"] == "Amazing Grace")
        arr = grace["arrangements"][0]
        assert len(arr["sections"]) == 3
        assert arr["sections"][0]["label"] == "Intro"
        assert arr["sections"][1]["label"] == "V1"
        assert arr["sections"][2]["label"] == "C1"

    @responses.activate
    def test_import_populates_sections_from_chord_chart_text(self, db_app):
        """Chord chart text is parsed into shorthand sections (strategy 2, after PDF)."""
        client, uid = db_app
        _mock_plan_items()
        _mock_arrangement_detail(
            "pco-song-1", "pco-arr-1",
            sequence=["Verse 1", "Chorus 1"],
            chord_chart="VERSE 1\nG D Em C\nCHORUS 1\nC G Am F\nBRIDGE 1\nEm D C",
        )
        _mock_arrangement_detail("pco-song-2", "pco-arr-2")

        client.post("/api/pco/import/plan-1", json={
            "serviceTypeId": "111", "date": "2026-03-08", "title": "Test",
        })

        songs = client.get("/api/songs").get_json()
        grace = next(s for s in songs if s["title"] == "Amazing Grace")
        arr = grace["arrangements"][0]
        labels = [s["label"] for s in arr["sections"]]
        assert "V1" in labels
        assert "C1" in labels
        assert "B1" in labels  # Only from chord chart, not in sequence

    @responses.activate
    def test_import_duplicate_returns_exists(self, db_app):
        """Second import of same plan returns exists=True instead of creating duplicate."""
        client, uid = db_app
        _mock_plan_items()
        _mock_arrangement_detail("pco-song-1", "pco-arr-1")
        _mock_arrangement_detail("pco-song-2", "pco-arr-2")

        resp1 = client.post("/api/pco/import/plan-1", json={
            "serviceTypeId": "111", "date": "2026-03-08", "title": "Test",
        })
        assert resp1.status_code == 201
        first_id = resp1.get_json()["setlistId"]

        # Second import without update flag — should return exists
        resp2 = client.post("/api/pco/import/plan-1", json={
            "serviceTypeId": "111", "date": "2026-03-08", "title": "Test",
        })
        assert resp2.status_code == 200
        data = resp2.get_json()
        assert data["exists"] is True
        assert data["setlistId"] == first_id

    @responses.activate
    def test_import_update_replaces_setlist(self, db_app):
        """Import with update=True re-fetches and replaces the existing setlist."""
        client, uid = db_app
        _mock_plan_items()
        _mock_arrangement_detail("pco-song-1", "pco-arr-1")
        _mock_arrangement_detail("pco-song-2", "pco-arr-2")

        resp1 = client.post("/api/pco/import/plan-1", json={
            "serviceTypeId": "111", "date": "2026-03-08", "title": "Test",
        })
        first_id = resp1.get_json()["setlistId"]

        # Re-import with update=True
        _mock_plan_items()
        _mock_arrangement_detail("pco-song-1", "pco-arr-1")
        _mock_arrangement_detail("pco-song-2", "pco-arr-2")

        resp2 = client.post("/api/pco/import/plan-1", json={
            "serviceTypeId": "111", "date": "2026-03-08", "title": "Updated Title",
            "update": True,
        })
        assert resp2.status_code == 201
        assert resp2.get_json()["setlistId"] == first_id

        setlists = client.get("/api/setlists").get_json()
        sl = next(s for s in setlists if s["id"] == first_id)
        assert sl["name"] == "Updated Title"

    @responses.activate
    def test_import_is_idempotent_songs(self, db_app):
        """Force-updating the same plan should not duplicate songs."""
        client, uid = db_app
        _mock_plan_items()
        _mock_arrangement_detail("pco-song-1", "pco-arr-1")
        _mock_arrangement_detail("pco-song-2", "pco-arr-2")

        client.post("/api/pco/import/plan-1", json={
            "serviceTypeId": "111", "date": "2026-03-08", "title": "Test",
        })

        _mock_plan_items()
        _mock_arrangement_detail("pco-song-1", "pco-arr-1")
        _mock_arrangement_detail("pco-song-2", "pco-arr-2")

        client.post("/api/pco/import/plan-1", json={
            "serviceTypeId": "111", "date": "2026-03-08", "title": "Test",
            "update": True,
        })

        songs = client.get("/api/songs").get_json()
        assert len(songs) == 2


@requires_db
class TestSongSearch:
    @responses.activate
    def test_search_songs(self, db_app):
        client, uid = db_app
        responses.get(
            f"{PCO_API}/services/v2/songs",
            json={"data": [
                {"id": "s1", "attributes": {"title": "Amazing Grace", "author": "Newton", "ccli_number": "1234", "arrangement_count": 2}},
            ]},
        )

        resp = client.get("/api/pco/songs?q=amazing")
        assert resp.status_code == 200
        songs = resp.get_json()
        assert len(songs) == 1
        assert songs[0]["title"] == "Amazing Grace"
        assert songs[0]["arrangementCount"] == 2

    def test_search_empty_query(self, db_app):
        client, uid = db_app
        resp = client.get("/api/pco/songs?q=")
        assert resp.status_code == 200
        assert resp.get_json() == []

    @responses.activate
    def test_list_arrangements(self, db_app):
        client, uid = db_app
        responses.get(
            f"{PCO_API}/services/v2/songs/s1/arrangements",
            json={"data": [
                {"id": "a1", "attributes": {"name": "Main", "chord_chart_key": "G", "bpm": 72, "chord_chart": None}},
                {"id": "a2", "attributes": {"name": "Acoustic", "chord_chart_key": "D", "bpm": 80, "chord_chart": "http://example.com/chart.pdf"}},
            ]},
        )

        resp = client.get("/api/pco/songs/s1/arrangements")
        assert resp.status_code == 200
        arrs = resp.get_json()
        assert len(arrs) == 2
        assert arrs[0]["name"] == "Main"
        assert arrs[1]["hasChordChart"] is True

    @responses.activate
    def test_import_single_arrangement(self, db_app):
        client, uid = db_app
        _mock_all_attachments()
        responses.get(
            f"{PCO_API}/services/v2/songs/s1",
            json={"data": {"attributes": {"title": "Test Song", "author": "Test"}}},
        )
        responses.get(
            f"{PCO_API}/services/v2/songs/s1/arrangements",
            json={"data": [
                {"id": "a1", "attributes": {"name": "Main", "chord_chart_key": "G", "bpm": 72}},
                {"id": "a2", "attributes": {"name": "Acoustic", "chord_chart_key": "D", "bpm": 80}},
            ]},
        )

        resp = client.post("/api/pco/songs/s1/import", json={"pcoArrangementId": "a1"})
        assert resp.status_code == 201
        data = resp.get_json()
        assert len(data["arrangementIds"]) == 1

        songs = client.get("/api/songs").get_json()
        assert len(songs) == 1
        assert len(songs[0]["arrangements"]) == 1
        assert songs[0]["arrangements"][0]["key"] == "G"

    @responses.activate
    def test_import_song_parses_chord_chart(self, db_app):
        """Single-song import should parse chord_chart text into shorthand sections."""
        client, uid = db_app
        _mock_all_attachments()
        responses.get(
            f"{PCO_API}/services/v2/songs/s1",
            json={"data": {"attributes": {"title": "Chart Song", "author": "Artist"}}},
        )
        responses.get(
            f"{PCO_API}/services/v2/songs/s1/arrangements",
            json={"data": [
                {
                    "id": "a1",
                    "attributes": {
                        "name": "Main", "chord_chart_key": "C", "bpm": 120,
                        "chord_chart": "INTRO\nG D\nVERSE 1\nEm C\nCHORUS 1\nG D Em C",
                        "sequence": [],
                    },
                },
            ]},
        )

        resp = client.post("/api/pco/songs/s1/import")
        assert resp.status_code == 201

        songs = client.get("/api/songs").get_json()
        arr = songs[0]["arrangements"][0]
        labels = [s["label"] for s in arr["sections"]]
        assert "Intro" in labels
        assert "V1" in labels
        assert "C1" in labels


@requires_db
class TestSectionCopy:
    """Test the 'arrange once, auto-populate' feature."""

    @responses.activate
    def test_sections_copied_on_reimport(self, db_app):
        client, uid = db_app
        _mock_all_attachments()

        responses.get(
            f"{PCO_API}/services/v2/songs/s1",
            json={"data": {"attributes": {"title": "Grace Song", "author": "Author"}}},
        )
        responses.get(
            f"{PCO_API}/services/v2/songs/s1/arrangements",
            json={"data": [
                {"id": "arr-x", "attributes": {"name": "Main", "chord_chart_key": "G", "bpm": 72}},
            ]},
        )
        resp = client.post("/api/pco/songs/s1/import")
        first_arr_id = resp.get_json()["arrangementIds"][0]

        client.post(f"/api/arrangements/{first_arr_id}/sections", json={
            "sections": [
                {"label": "V1", "energy": "down", "notes": "quiet"},
                {"label": "C1", "energy": "up", "notes": "big"},
            ],
        })

        client.delete(f"/api/arrangements/{first_arr_id}")

        responses.get(
            f"{PCO_API}/services/v2/songs/s1",
            json={"data": {"attributes": {"title": "Grace Song", "author": "Author"}}},
        )
        responses.get(
            f"{PCO_API}/services/v2/songs/s1/arrangements",
            json={"data": [
                {"id": "arr-x", "attributes": {"name": "Main", "chord_chart_key": "G", "bpm": 72}},
            ]},
        )
        resp2 = client.post("/api/pco/songs/s1/import")
        songs = client.get("/api/songs").get_json()
        song = next(s for s in songs if s["title"] == "Grace Song")
        assert len(song["arrangements"][0]["sections"]) == 0

    @responses.activate
    def test_sections_copied_when_donor_exists(self, db_app):
        """If an arrangement with the same pco_arrangement_id exists and has sections, copy them."""
        client, uid = db_app
        _mock_all_attachments()

        responses.get(
            f"{PCO_API}/services/v2/songs/s2",
            json={"data": {"attributes": {"title": "Copy Test", "author": "Author"}}},
        )
        responses.get(
            f"{PCO_API}/services/v2/songs/s2/arrangements",
            json={"data": [
                {"id": "arr-copy", "attributes": {"name": "Main", "chord_chart_key": "C", "bpm": 90}},
            ]},
        )
        resp = client.post("/api/pco/songs/s2/import")
        arr_id = resp.get_json()["arrangementIds"][0]

        client.post(f"/api/arrangements/{arr_id}/sections", json={
            "sections": [
                {"label": "Intro", "energy": "steady", "notes": "keys only"},
                {"label": "V1", "energy": "build", "notes": "add guitar"},
            ],
        })

        songs = client.get("/api/songs").get_json()
        song = next(s for s in songs if s["title"] == "Copy Test")
        assert len(song["arrangements"][0]["sections"]) == 2
        assert song["arrangements"][0]["sections"][0]["notes"] == "keys only"


@requires_db
class TestShorthandLabels:
    """Test PCO sequence label to shorthand conversion."""

    def test_shorthand_conversion(self, db_app):
        from preppy.pco import _shorthand_label
        assert _shorthand_label("Verse 1") == "V1"
        assert _shorthand_label("Verse 2") == "V2"
        assert _shorthand_label("Chorus 1") == "C1"
        assert _shorthand_label("Chorus 2") == "C2"
        assert _shorthand_label("Bridge 1") == "B1"
        assert _shorthand_label("Pre-Chorus 1") == "Pre 1"
        assert _shorthand_label("Intro") == "Intro"
        assert _shorthand_label("Outro") == "Outro"
        assert _shorthand_label("Instrumental") == "Instr"
        assert _shorthand_label("Tag") == "Tag"
        assert _shorthand_label("Turn") == "Turn"

    def test_shorthand_passthrough(self, db_app):
        """Unknown labels should pass through unchanged."""
        from preppy.pco import _shorthand_label
        assert _shorthand_label("Something Custom") == "Something Custom"
        assert _shorthand_label("V1") == "V1"


@requires_db
class TestChordChartParsing:
    """Test parsing chord chart text into sections."""

    def test_parse_chord_chart_text(self, db_app):
        from preppy.pco import _sections_from_chord_chart
        sections = _sections_from_chord_chart(
            "VERSE 1\nG D Em C\nCHORUS 1\nC G Am F\nBRIDGE 1\nEm D"
        )
        labels = [s["label"] for s in sections]
        assert "V1" in labels
        assert "C1" in labels
        assert "B1" in labels

    def test_parse_empty_chart(self, db_app):
        from preppy.pco import _sections_from_chord_chart
        assert _sections_from_chord_chart("") == []
        assert _sections_from_chord_chart("   ") == []

    def test_parse_chart_chords_only(self, db_app):
        """A chart with only chord lines and no section headers yields no sections."""
        from preppy.pco import _sections_from_chord_chart
        sections = _sections_from_chord_chart("G D Em C\nAm F G C")
        assert sections == []


@requires_db
class TestSync:
    @responses.activate
    def test_sync_updates_setlist(self, db_app):
        client, uid = db_app

        _mock_plan_items()
        _mock_arrangement_detail("pco-song-1", "pco-arr-1")
        _mock_arrangement_detail("pco-song-2", "pco-arr-2")

        import_resp = client.post("/api/pco/import/plan-1", json={
            "serviceTypeId": "111", "date": "2026-03-08", "title": "Test",
        })
        sl_id = import_resp.get_json()["setlistId"]

        responses.get(
            f"{PCO_API}/services/v2/service_types/111/plans/plan-1/items",
            json={
                "data": [
                    {
                        "id": "item-2",
                        "attributes": {"item_type": "song", "title": "Amazing Grace", "sequence": 0},
                        "relationships": {
                            "song": {"data": {"id": "pco-song-1"}},
                            "arrangement": {"data": {"id": "pco-arr-1"}},
                        },
                    },
                    {
                        "id": "item-4",
                        "attributes": {"item_type": "song", "title": "New Song", "sequence": 1},
                        "relationships": {
                            "song": {"data": {"id": "pco-song-3"}},
                            "arrangement": {"data": {"id": "pco-arr-3"}},
                        },
                    },
                ],
                "included": [
                    {"id": "pco-song-1", "type": "Song", "attributes": {"title": "Amazing Grace", "author": "Newton"}},
                    {"id": "pco-song-3", "type": "Song", "attributes": {"title": "New Song", "author": "New Author"}},
                    {"id": "pco-arr-1", "type": "Arrangement", "attributes": {"name": "Main", "chord_chart_key": "G", "bpm": 72}},
                    {"id": "pco-arr-3", "type": "Arrangement", "attributes": {"name": "Main", "chord_chart_key": "E", "bpm": 110}},
                ],
            },
        )
        _mock_arrangement_detail("pco-song-3", "pco-arr-3", ["Verse 1", "Chorus 1"])

        sync_resp = client.post("/api/pco/plans/plan-1/sync", json={
            "serviceTypeId": "111", "setlistId": sl_id,
        })
        assert sync_resp.status_code == 200
        changes = sync_resp.get_json()["changes"]
        assert changes["added"] == 1
        assert changes["removed"] == 1

        setlists = client.get("/api/setlists").get_json()
        sl = next(s for s in setlists if s["id"] == sl_id)
        song_items = [i for i in sl["items"] if i["itemType"] == "song"]
        assert len(song_items) == 2
        titles = {i["title"] for i in song_items}
        assert "Amazing Grace" in titles
        assert "New Song" in titles
        assert "How Great Thou Art" not in titles


@requires_db
class TestUploadPrepSheet:
    @responses.activate
    def test_upload_two_step(self, db_app):
        client, uid = db_app

        _mock_plan_items()
        _mock_arrangement_detail("pco-song-1", "pco-arr-1")
        _mock_arrangement_detail("pco-song-2", "pco-arr-2")
        client.post("/api/pco/import/plan-1", json={
            "serviceTypeId": "111", "date": "2026-03-08", "title": "Test",
        })

        responses.post(
            PCO_UPLOAD,
            json={"data": {"id": "upload-uuid-123"}},
        )
        responses.post(
            f"{PCO_API}/services/v2/service_types/111/plans/plan-1/attachments",
            json={"data": {"id": "attachment-1", "attributes": {"url": "https://pco.test/file.pdf"}}},
        )

        import io
        data = {
            "file": (io.BytesIO(b"%PDF-fake"), "Prep Sheet.pdf", "application/pdf"),
            "serviceTypeId": "111",
            "filename": "Prep Sheet.pdf",
        }
        resp = client.post(
            "/api/pco/plans/plan-1/upload-prep-sheet",
            data=data,
            content_type="multipart/form-data",
        )
        assert resp.status_code == 200
        result = resp.get_json()
        assert result["attachmentId"] == "attachment-1"
        assert "pco.test" in result["url"]

    @responses.activate
    def test_upload_with_generated_pdf(self, db_app):
        """End-to-end: generate a real PDF via /api/export-pdf, then upload it."""
        client, uid = db_app

        _mock_plan_items()
        _mock_arrangement_detail("pco-song-1", "pco-arr-1")
        _mock_arrangement_detail("pco-song-2", "pco-arr-2")
        client.post("/api/pco/import/plan-1", json={
            "serviceTypeId": "111", "date": "2026-03-08", "title": "Test",
        })

        pdf_resp = client.post("/api/export-pdf", json={
            "lines": ["Prep Sheet March 8, 2026", "Amazing Grace [G] - 72BPM"],
            "filename": "Prep Sheet.pdf",
        })
        assert pdf_resp.status_code == 200
        assert pdf_resp.data[:5] == b"%PDF-"

        responses.post(
            PCO_UPLOAD,
            json={"data": {"id": "upload-uuid-789"}},
        )
        responses.post(
            f"{PCO_API}/services/v2/service_types/111/plans/plan-1/attachments",
            json={"data": {"id": "attachment-3", "attributes": {"url": "https://pco.test/prep.pdf"}}},
        )

        import io
        data = {
            "file": (io.BytesIO(pdf_resp.data), "Prep Sheet.pdf", "application/pdf"),
            "serviceTypeId": "111",
            "filename": "Prep Sheet.pdf",
        }
        resp = client.post(
            "/api/pco/plans/plan-1/upload-prep-sheet",
            data=data,
            content_type="multipart/form-data",
        )
        assert resp.status_code == 200
        result = resp.get_json()
        assert result["attachmentId"] == "attachment-3"

    def test_upload_requires_file(self, db_app):
        client, uid = db_app
        resp = client.post(
            "/api/pco/plans/plan-1/upload-prep-sheet",
            data={"serviceTypeId": "111"},
            content_type="multipart/form-data",
        )
        assert resp.status_code == 400

    def test_upload_requires_service_type(self, db_app):
        client, uid = db_app
        import io
        resp = client.post(
            "/api/pco/plans/plan-1/upload-prep-sheet",
            data={"file": (io.BytesIO(b"data"), "test.pdf")},
            content_type="multipart/form-data",
        )
        assert resp.status_code == 400
