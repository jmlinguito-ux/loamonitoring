"""Tests for safe retry behavior on Dataverse record updates."""

import unittest
from unittest.mock import MagicMock, patch

import requests

from dataverse import DataverseClient


class TestDataversePatchRetry(unittest.TestCase):
    def make_client(self):
        client = object.__new__(DataverseClient)
        client.api_root = 'https://example.crm.dynamics.com/api/data/v9.2/'
        client._headers = MagicMock(return_value={'Authorization': 'Bearer test'})
        client._session = MagicMock()
        return client

    @patch('dataverse.time.sleep')
    def test_patch_retries_transient_connection_abort(self, sleep):
        client = self.make_client()
        success = MagicMock(status_code=204)
        client._session.patch.side_effect = [
            requests.exceptions.ConnectionError('Connection aborted: RemoteDisconnected'),
            success,
        ]

        status = client.patch('cr63f_loamonitorings', 'record-1', {'cr63f_balance': 123.45})

        self.assertEqual(status, 204)
        self.assertEqual(client._session.patch.call_count, 2)
        sleep.assert_called_once_with(0.5)

    @patch('dataverse.time.sleep')
    def test_patch_stops_after_three_transient_failures(self, sleep):
        client = self.make_client()
        client._session.patch.side_effect = requests.exceptions.ConnectionError('RemoteDisconnected')

        with self.assertRaises(requests.exceptions.ConnectionError):
            client.patch('cr63f_loamonitorings', 'record-1', {'cr63f_balance': 123.45})

        self.assertEqual(client._session.patch.call_count, 3)
        self.assertEqual(sleep.call_count, 2)


if __name__ == '__main__':
    unittest.main()
