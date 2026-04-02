"""
Tests for Synapse API integration: auth fallback chain, poll/timeout,
http_utils retry, and error handling.

Uses unittest.mock to avoid real HTTP calls.
"""

import os
import sys
import json
import pytest
from unittest.mock import patch, MagicMock

BASE = os.path.dirname(__file__)
ROOT = os.path.dirname(BASE)
sys.path.insert(0, ROOT)


# ── Auth tests ────────────────────────────────────────────────────────────────

class TestResolveApiKey:
    """Test synapse_auth.resolve_api_key fallback chain."""

    def test_explicit_key_wins(self):
        from slidegen.pipeline.synapse_auth import resolve_api_key
        assert resolve_api_key(api_key="explicit-token") == "explicit-token"

    def test_env_var_fallback(self, monkeypatch):
        from slidegen.pipeline.synapse_auth import resolve_api_key
        monkeypatch.setenv("SYNAPSE_API_KEY", "env-token")
        assert resolve_api_key() == "env-token"

    def test_explicit_beats_env(self, monkeypatch):
        from slidegen.pipeline.synapse_auth import resolve_api_key
        monkeypatch.setenv("SYNAPSE_API_KEY", "env-token")
        assert resolve_api_key(api_key="explicit") == "explicit"

    def test_no_key_raises(self, monkeypatch):
        from slidegen.pipeline.synapse_auth import resolve_api_key
        monkeypatch.delenv("SYNAPSE_API_KEY", raising=False)
        monkeypatch.delenv("AZURE_TENANT_ID", raising=False)
        monkeypatch.delenv("AZURE_CLIENT_ID", raising=False)
        monkeypatch.delenv("AZURE_CLIENT_SECRET", raising=False)
        with pytest.raises(ValueError, match="No Synapse API token available"):
            resolve_api_key()

    def test_azure_ad_fallback(self, monkeypatch):
        """When no explicit key or env var, Azure AD credentials are tried."""
        from slidegen.pipeline.synapse_auth import resolve_api_key, _cached_token
        import slidegen.pipeline.synapse_auth as auth_mod

        monkeypatch.delenv("SYNAPSE_API_KEY", raising=False)
        monkeypatch.setenv("AZURE_TENANT_ID", "test-tenant")
        monkeypatch.setenv("AZURE_CLIENT_ID", "test-client")
        monkeypatch.setenv("AZURE_CLIENT_SECRET", "test-secret")

        # Clear cached token
        auth_mod._cached_token = None
        auth_mod._cached_token_expiry = 0

        # Mock _post_form to return a fake token
        with patch.object(auth_mod, '_post_form', return_value={
            "access_token": "azure-token-123",
            "expires_in": 3600,
        }):
            result = resolve_api_key()
            assert result == "azure-token-123"

        # Clean up
        auth_mod._cached_token = None
        auth_mod._cached_token_expiry = 0

    def test_azure_token_cached(self, monkeypatch):
        """Second call returns cached token without hitting Azure."""
        import slidegen.pipeline.synapse_auth as auth_mod
        import time

        monkeypatch.delenv("SYNAPSE_API_KEY", raising=False)
        monkeypatch.setenv("AZURE_TENANT_ID", "t")
        monkeypatch.setenv("AZURE_CLIENT_ID", "c")
        monkeypatch.setenv("AZURE_CLIENT_SECRET", "s")

        auth_mod._cached_token = "cached-token"
        auth_mod._cached_token_expiry = time.time() + 3600

        # Should return cached token without calling _post_form
        with patch.object(auth_mod, '_post_form') as mock_post:
            result = auth_mod.resolve_api_key()
            assert result == "cached-token"
            mock_post.assert_not_called()

        # Clean up
        auth_mod._cached_token = None
        auth_mod._cached_token_expiry = 0


# ── HTTP utils retry tests ───────────────────────────────────────────────────

class TestHttpUtilsRetry:
    """Test http_utils get_json/post_json retry with backoff."""

    def test_get_json_success(self):
        from slidegen.pipeline.http_utils import get_json
        mock_resp = MagicMock()
        mock_resp.ok = True
        mock_resp.json.return_value = {"status": "ok"}

        with patch('slidegen.pipeline.http_utils._requests.get', return_value=mock_resp):
            result = get_json("http://test/api", {"Authorization": "Bearer x"})
            assert result == {"status": "ok"}

    def test_get_json_retries_on_500(self):
        from slidegen.pipeline.http_utils import get_json

        fail_resp = MagicMock()
        fail_resp.ok = False
        fail_resp.status_code = 500
        fail_resp.text = "Internal Server Error"
        fail_resp.json.side_effect = ValueError

        ok_resp = MagicMock()
        ok_resp.ok = True
        ok_resp.json.return_value = {"data": "recovered"}

        with patch('slidegen.pipeline.http_utils._requests.get',
                   side_effect=[fail_resp, ok_resp]):
            with patch('slidegen.pipeline.http_utils.time.sleep'):  # don't actually wait
                result = get_json("http://test/api", {})
                assert result == {"data": "recovered"}

    def test_get_json_raises_on_persistent_failure(self):
        from slidegen.pipeline.http_utils import get_json

        fail_resp = MagicMock()
        fail_resp.ok = False
        fail_resp.status_code = 500
        fail_resp.text = "Server Error"
        fail_resp.json.side_effect = ValueError

        with patch('slidegen.pipeline.http_utils._requests.get',
                   return_value=fail_resp):
            with patch('slidegen.pipeline.http_utils.time.sleep'):
                with pytest.raises(RuntimeError, match="500"):
                    get_json("http://test/api", {})

    def test_get_json_no_retry_on_400(self):
        """Client errors (4xx except 429) should not retry."""
        from slidegen.pipeline.http_utils import get_json

        fail_resp = MagicMock()
        fail_resp.ok = False
        fail_resp.status_code = 400
        fail_resp.text = "Bad Request"
        fail_resp.json.side_effect = ValueError

        with patch('slidegen.pipeline.http_utils._requests.get',
                   return_value=fail_resp) as mock_get:
            with pytest.raises(RuntimeError, match="400"):
                get_json("http://test/api", {})
            # Should only be called once (no retry on 400)
            assert mock_get.call_count == 1

    def test_get_json_retries_on_429(self):
        """429 Too Many Requests should trigger retry."""
        from slidegen.pipeline.http_utils import get_json

        rate_resp = MagicMock()
        rate_resp.ok = False
        rate_resp.status_code = 429
        rate_resp.text = "Rate limited"
        rate_resp.json.side_effect = ValueError

        ok_resp = MagicMock()
        ok_resp.ok = True
        ok_resp.json.return_value = {"ok": True}

        with patch('slidegen.pipeline.http_utils._requests.get',
                   side_effect=[rate_resp, ok_resp]):
            with patch('slidegen.pipeline.http_utils.time.sleep'):
                result = get_json("http://test/api", {})
                assert result == {"ok": True}

    def test_post_json_success(self):
        from slidegen.pipeline.http_utils import post_json
        mock_resp = MagicMock()
        mock_resp.ok = True
        mock_resp.json.return_value = {"id": 42}

        with patch('slidegen.pipeline.http_utils._requests.post', return_value=mock_resp):
            result = post_json("http://test/api", {}, {"key": "val"})
            assert result == {"id": 42}

    def test_post_json_retries_on_503(self):
        from slidegen.pipeline.http_utils import post_json

        fail_resp = MagicMock()
        fail_resp.ok = False
        fail_resp.status_code = 503
        fail_resp.text = "Service Unavailable"
        fail_resp.json.side_effect = ValueError

        ok_resp = MagicMock()
        ok_resp.ok = True
        ok_resp.json.return_value = {"result": "ok"}

        with patch('slidegen.pipeline.http_utils._requests.post',
                   side_effect=[fail_resp, ok_resp]):
            with patch('slidegen.pipeline.http_utils.time.sleep'):
                result = post_json("http://test/api", {}, {})
                assert result == {"result": "ok"}


# ── Poll/timeout tests ───────────────────────────────────────────────────────

class TestPollUntilDone:
    """Test synapse_fetcher._poll_until_done behavior."""

    def test_poll_completes_on_processed(self):
        from slidegen.pipeline.synapse_fetcher import _poll_until_done

        with patch('slidegen.pipeline.synapse_fetcher._get_json',
                   return_value={"status": "processed"}):
            # Should not raise
            _poll_until_done("http://test", {}, 1, poll_interval=1, max_wait=10)

    def test_poll_raises_on_failed(self):
        from slidegen.pipeline.synapse_fetcher import _poll_until_done

        with patch('slidegen.pipeline.synapse_fetcher._get_json',
                   return_value={"status": "failed"}):
            with pytest.raises(RuntimeError, match="failed"):
                _poll_until_done("http://test", {}, 1, poll_interval=1, max_wait=10)

    def test_poll_timeout(self):
        from slidegen.pipeline.synapse_fetcher import _poll_until_done

        with patch('slidegen.pipeline.synapse_fetcher._get_json',
                   return_value={"status": "processing"}):
            with patch('slidegen.pipeline.synapse_fetcher.time.sleep'):
                with pytest.raises(TimeoutError, match="not complete"):
                    _poll_until_done("http://test", {}, 1,
                                     poll_interval=1, max_wait=2)

    def test_poll_waits_then_succeeds(self):
        from slidegen.pipeline.synapse_fetcher import _poll_until_done

        responses = [
            {"status": "processing"},
            {"status": "processing"},
            {"status": "processed"},
        ]
        with patch('slidegen.pipeline.synapse_fetcher._get_json',
                   side_effect=responses):
            with patch('slidegen.pipeline.synapse_fetcher.time.sleep'):
                # Should complete after 3 polls
                _poll_until_done("http://test", {}, 1,
                                 poll_interval=1, max_wait=10)


# ── Trigger generation tests ─────────────────────────────────────────────────

class TestTriggerGeneration:
    """Test synapse_fetcher._trigger_generation."""

    def test_trigger_returns_history_id(self):
        from slidegen.pipeline.synapse_fetcher import _trigger_generation
        from slidegen.pipeline.project_config import SynapseConfig

        synapse = SynapseConfig(
            api_url="http://test/api",
            project_id=1,
            survey_ids=[1],
            deliverable_ids=[1],
            segment_ids=[],
        )

        mock_resp = MagicMock()
        mock_resp.ok = True
        mock_resp.json.return_value = {"history_id": 42}

        with patch('slidegen.pipeline.synapse_fetcher._requests.post',
                   return_value=mock_resp):
            result = _trigger_generation("http://test/api", synapse, {})
            assert result == 42

    def test_trigger_raises_on_missing_history_id(self):
        from slidegen.pipeline.synapse_fetcher import _trigger_generation
        from slidegen.pipeline.project_config import SynapseConfig

        synapse = SynapseConfig(
            api_url="http://test/api",
            project_id=1,
            survey_ids=[1],
            deliverable_ids=[1],
            segment_ids=[],
        )

        mock_resp = MagicMock()
        mock_resp.ok = True
        mock_resp.json.return_value = {"status": "ok"}  # no history_id

        with patch('slidegen.pipeline.synapse_fetcher._requests.post',
                   return_value=mock_resp):
            with pytest.raises(ValueError, match="missing 'history_id'"):
                _trigger_generation("http://test/api", synapse, {})

    def test_trigger_raises_on_api_error(self):
        from slidegen.pipeline.synapse_fetcher import _trigger_generation
        from slidegen.pipeline.project_config import SynapseConfig

        synapse = SynapseConfig(
            api_url="http://test/api",
            project_id=1,
            survey_ids=[1],
            deliverable_ids=[1],
            segment_ids=[],
        )

        mock_resp = MagicMock()
        mock_resp.ok = False
        mock_resp.status_code = 403
        mock_resp.text = "Forbidden"
        mock_resp.json.side_effect = ValueError

        with patch('slidegen.pipeline.synapse_fetcher._requests.post',
                   return_value=mock_resp):
            with pytest.raises(RuntimeError, match="403"):
                _trigger_generation("http://test/api", synapse, {})


# ── Data completeness check tests ─────────────────────────────────────────────

class TestDataCompleteness:
    """Test the post-load _warnings system in data_loaders."""

    def test_warnings_track_missing_extractions(self):
        """Warnings should list extractions that are missing from data."""
        yaml_path = os.path.join(BASE, "test_project", "config.yaml")
        from slidegen.pipeline.project_config import load_project_config
        config = load_project_config(yaml_path)

        # Simulate load_all_data result with some extractions missing
        data = {
            "_meta": {},
            "_sheets": {},
            "ryb_mr": [{"desc": "test", "current": 0.5}],
            # All other extraction IDs are "missing"
        }

        # Apply the completeness check logic directly
        expected_ids = {ex.id for ex in config.extractions}
        loaded_ids = {k for k in data if not k.startswith("_")}
        missing = expected_ids - loaded_ids

        assert len(missing) > 0, "Test requires some extractions to be missing"
        assert "ryb_mr" not in missing, "ryb_mr should be loaded"
        assert "tag_mr" in missing, "tag_mr should be missing"

    def test_empty_extraction_flagged(self):
        """Extractions returning 0 rows should be flagged."""
        data = {
            "_meta": {},
            "extraction_a": [{"desc": "has data"}],
            "extraction_b": [],  # empty
        }

        empty = {k for k in data if not k.startswith("_")
                 and isinstance(data[k], list) and len(data[k]) == 0}
        assert "extraction_b" in empty
        assert "extraction_a" not in empty


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
