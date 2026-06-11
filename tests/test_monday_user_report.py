"""
Tests for monday_user_report.py

Run with:
    pip install pytest requests python-dotenv
    pytest test_monday_user_report.py -v
"""

import csv
import io
import os
import sys
from unittest.mock import MagicMock, patch

import pytest

# Make the report module importable without a token in the environment
os.environ.setdefault("MONDAY_API_TOKEN", "test-token")

import monday_user_report as report


# ── Fixtures ──────────────────────────────────────────────────

def make_member(**overrides) -> dict:
    """Return a base member dict with optional field overrides."""
    base = {
        "id": "1",
        "name": "Alice Smith",
        "email": "alice@example.com",
        "is_admin": False,
        "is_guest": False,
        "is_view_only": False,
        "enabled": True,
        "created_at": "2024-01-15",
        "last_activity": "2024-06-01",
        "teams": [{"name": "Engineering"}],
    }
    base.update(overrides)
    return base


# ── get_role ──────────────────────────────────────────────────

class TestGetRole:
    def test_admin(self):
        assert report.get_role(make_member(is_admin=True)) == "Admin"

    def test_guest(self):
        assert report.get_role(make_member(is_guest=True)) == "Guest"

    def test_viewer(self):
        assert report.get_role(make_member(is_view_only=True)) == "Viewer"

    def test_regular_member(self):
        assert report.get_role(make_member()) == "Member"

    def test_admin_takes_priority_over_guest(self):
        # is_admin should win if both flags are somehow set
        assert report.get_role(make_member(is_admin=True, is_guest=True)) == "Admin"

    def test_empty_member(self):
        assert report.get_role({}) == "Member"


# ── get_status ────────────────────────────────────────────────

class TestGetStatus:
    def test_active(self):
        assert report.get_status(make_member(enabled=True)) == "Active"

    def test_inactive(self):
        assert report.get_status(make_member(enabled=False)) == "Inactive"

    def test_missing_enabled_defaults_inactive(self):
        assert report.get_status({}) == "Inactive"


# ── get_teams ─────────────────────────────────────────────────

class TestGetTeams:
    def test_single_team(self):
        m = make_member(teams=[{"name": "Engineering"}])
        assert report.get_teams(m) == "Engineering"

    def test_multiple_teams(self):
        m = make_member(teams=[{"name": "Engineering"}, {"name": "Design"}])
        assert report.get_teams(m) == "Engineering; Design"

    def test_no_teams_empty_list(self):
        assert report.get_teams(make_member(teams=[])) == "No Teams"

    def test_no_teams_missing_key(self):
        assert report.get_teams({}) == "No Teams"

    def test_no_teams_none_value(self):
        assert report.get_teams(make_member(teams=None)) == "No Teams"


# ── build_row ─────────────────────────────────────────────────

class TestBuildRow:
    WS_NAME = "Main Workspace"
    WS_URL  = "https://coral.monday.com/workspaces/42"

    def test_row_length(self):
        row = report.build_row(make_member(), self.WS_NAME, self.WS_URL)
        assert len(row) == 10

    def test_column_order(self):
        member = make_member(
            name="Bob Jones",
            email="bob@example.com",
            is_admin=True,
            enabled=True,
            created_at="2023-05-01",
            last_activity="2024-07-10",
            teams=[{"name": "Ops"}],
        )
        row = report.build_row(member, self.WS_NAME, self.WS_URL)
        workspace, name, email, role, status, teams, joined, last_active, twofa, url = row

        assert workspace    == self.WS_NAME
        assert name         == "Bob Jones"
        assert email        == "bob@example.com"
        assert role         == "Admin"
        assert status       == "Active"
        assert teams        == "Ops"
        assert joined       == "2023-05-01"
        assert last_active  == "2024-07-10"
        assert twofa        == "Disabled"
        assert url          == self.WS_URL

    def test_missing_dates_use_defaults(self):
        member = make_member(created_at=None, last_activity=None)
        row    = report.build_row(member, self.WS_NAME, self.WS_URL)
        assert row[6] == ""                  # Joined
        assert row[7] == "Never logged in"   # Last Active

    def test_2fa_always_disabled(self):
        row = report.build_row(make_member(), self.WS_NAME, self.WS_URL)
        assert row[8] == "Disabled"

    def test_workspace_url_in_last_column(self):
        row = report.build_row(make_member(), self.WS_NAME, self.WS_URL)
        assert row[9] == self.WS_URL


# ── gql (HTTP layer) ─────────────────────────────────────────

class TestGql:
    TOKEN = "test-token"
    QUERY = "query { workspaces { id } }"

    def _mock_response(self, json_data: dict, status_code: int = 200) -> MagicMock:
        mock = MagicMock()
        mock.status_code = status_code
        mock.json.return_value = json_data
        mock.raise_for_status = MagicMock()
        return mock

    @patch("monday_user_report.requests.post")
    def test_returns_parsed_json(self, mock_post):
        payload = {"data": {"workspaces": []}}
        mock_post.return_value = self._mock_response(payload)

        result = report.gql(self.QUERY, self.TOKEN)
        assert result == payload

    @patch("monday_user_report.requests.post")
    def test_sends_correct_headers(self, mock_post):
        mock_post.return_value = self._mock_response({"data": {}})
        report.gql(self.QUERY, self.TOKEN)

        _, kwargs = mock_post.call_args
        headers = kwargs["headers"]
        assert headers["Authorization"] == self.TOKEN
        assert headers["API-Version"]   == "2024-01"
        assert headers["Content-Type"]  == "application/json"

    @patch("monday_user_report.requests.post")
    def test_sends_query_in_body(self, mock_post):
        mock_post.return_value = self._mock_response({"data": {}})
        report.gql(self.QUERY, self.TOKEN)

        _, kwargs = mock_post.call_args
        assert kwargs["json"]["query"] == self.QUERY

    @patch("monday_user_report.requests.post")
    def test_raises_on_graphql_errors(self, mock_post):
        mock_post.return_value = self._mock_response(
            {"errors": [{"message": "Unauthorized"}]}
        )
        with pytest.raises(RuntimeError, match="GraphQL error"):
            report.gql(self.QUERY, self.TOKEN)

    @patch("monday_user_report.requests.post")
    def test_raises_on_http_error(self, mock_post):
        mock = self._mock_response({}, status_code=500)
        mock.raise_for_status.side_effect = Exception("500 Server Error")
        mock_post.return_value = mock

        with pytest.raises(Exception, match="500 Server Error"):
            report.gql(self.QUERY, self.TOKEN)


# ── main (integration-style) ──────────────────────────────────

class TestMain:
    """End-to-end test of main() with all I/O mocked."""

    WORKSPACES_RESPONSE = {
        "data": {
            "workspaces": [
                {"id": "10", "name": "Alpha"},
                {"id": "20", "name": "Beta"},
            ]
        }
    }

    MEMBERS_ALPHA = {
        "data": {
            "workspaces": [{
                "members": [
                    make_member(name="Alice", email="alice@a.com", is_admin=True,
                                teams=[{"name": "Eng"}]),
                    make_member(name="Bob",   email="bob@a.com",   is_guest=True,
                                enabled=False, teams=[]),
                ]
            }]
        }
    }

    MEMBERS_BETA = {
        "data": {
            "workspaces": [{"members": []}]   # empty workspace
        }
    }

    @patch("monday_user_report.time.sleep")
    @patch("monday_user_report.requests.post")
    def test_csv_output(self, mock_post, mock_sleep, tmp_path, monkeypatch):
        monkeypatch.setenv("MONDAY_API_TOKEN", "fake-token")
        monkeypatch.chdir(tmp_path)

        responses = [
            self._make_resp(self.WORKSPACES_RESPONSE),
            self._make_resp(self.MEMBERS_ALPHA),
            self._make_resp(self.MEMBERS_BETA),
        ]
        mock_post.side_effect = responses

        report.main()

        # Find the generated CSV
        csvfiles = list(tmp_path.glob("coral_user_report_*.csv"))
        assert len(csvfiles) == 1

        rows = list(csv.reader(csvfiles[0].open(encoding="utf-8")))
        header, *data_rows = rows

        assert header[0] == "Workspace"
        assert header[3] == "User Role"

        # Alpha has 2 members, Beta has 0 → 2 data rows
        assert len(data_rows) == 2

        alice_row = data_rows[0]
        assert alice_row[0] == "Alpha"
        assert alice_row[2] == "alice@a.com"
        assert alice_row[3] == "Admin"
        assert alice_row[4] == "Active"
        assert alice_row[5] == "Eng"

        bob_row = data_rows[1]
        assert bob_row[3] == "Guest"
        assert bob_row[4] == "Inactive"
        assert bob_row[5] == "No Teams"

    @patch("monday_user_report.time.sleep")
    @patch("monday_user_report.requests.post")
    def test_sleep_called_per_workspace(self, mock_post, mock_sleep, tmp_path, monkeypatch):
        monkeypatch.setenv("MONDAY_API_TOKEN", "fake-token")
        monkeypatch.chdir(tmp_path)

        mock_post.side_effect = [
            self._make_resp(self.WORKSPACES_RESPONSE),
            self._make_resp(self.MEMBERS_ALPHA),
            self._make_resp(self.MEMBERS_BETA),
        ]

        report.main()
        # sleep is called once per workspace that has members;
        # Beta has no members so the loop hits `continue` before sleep
        assert mock_sleep.call_count == 1

    def _make_resp(self, data: dict) -> MagicMock:
        m = MagicMock()
        m.json.return_value = data
        m.raise_for_status = MagicMock()
        return m
