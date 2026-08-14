import base64
import json
import threading
from http.client import HTTPConnection
from tempfile import TemporaryDirectory
from pathlib import Path
import unittest

from lucas_cloud.server import make_server
from lucas_cloud.service import LucasCloudService


class LucasCloudPlatformTests(unittest.TestCase):
    def test_workspace_photo_upload_pending_and_status_lifecycle(self) -> None:
        with TemporaryDirectory() as tmp:
            service = LucasCloudService(Path(tmp))
            created = service.create_workspace("Danny LUCAS", "Danny LUCAS")
            token = created["token"]

            upload = service.upload_mobile_photos(
                "danny-lucas",
                token,
                {
                    "uploader": "Danny",
                    "assigned_person": "Danny",
                    "images": [
                        {
                            "name": "front.jpg",
                            "image": "data:image/jpeg;base64," + base64.b64encode(b"jpg-bytes").decode("ascii"),
                        }
                    ],
                },
            )

            self.assertTrue(upload["ok"])
            self.assertEqual(upload["saved"], 1)
            photo = upload["photos"][0]
            self.assertEqual(photo["status"], "pending_scan")
            self.assertEqual(photo["assigned_person"], "Danny")
            stored = Path(tmp) / "objects" / photo["storage_path"]
            self.assertTrue(stored.exists())
            self.assertEqual(stored.read_bytes(), b"jpg-bytes")

            pending = service.pending_photos("danny-lucas", token)
            self.assertEqual([item["id"] for item in pending["photos"]], [photo["id"]])

            linked = service.mark_photo_status(
                "danny-lucas",
                token,
                photo["id"],
                {"status": "linked", "linked_inventory_id": "inventory-1", "actor": "desktop"},
            )
            self.assertEqual(linked["photo"]["status"], "linked")
            self.assertEqual(linked["photo"]["linked_inventory_id"], "inventory-1")
            self.assertEqual(service.pending_photos("danny-lucas", token)["photos"], [])

    def test_workspace_token_is_required(self) -> None:
        with TemporaryDirectory() as tmp:
            service = LucasCloudService(Path(tmp))
            created = service.create_workspace("team-lucas", "Team LUCAS")

            with self.assertRaises(PermissionError):
                service.upload_mobile_photos(
                    "team-lucas",
                    "bad-token",
                    {"images": [{"name": "front.jpg", "image": "data:image/jpeg;base64,anBn"}]},
                )
            with self.assertRaises(PermissionError):
                service.pending_photos("team-lucas", "bad-token")
            self.assertTrue(created["token"])

    def test_inventory_upsert_is_workspace_scoped(self) -> None:
        with TemporaryDirectory() as tmp:
            service = LucasCloudService(Path(tmp))
            danny = service.create_workspace("danny", "Danny")["token"]
            team = service.create_workspace("team-lucas", "Team LUCAS")["token"]

            first = service.upsert_inventory("danny", danny, {"external_key": "79709194", "card_title": "Danny Card"})
            second = service.upsert_inventory("team-lucas", team, {"external_key": "79709194", "card_title": "Team Card"})

            self.assertNotEqual(first["item"]["workspace_id"], second["item"]["workspace_id"])
            self.assertEqual(first["item"]["card_title"], "Danny Card")
            self.assertEqual(second["item"]["card_title"], "Team Card")

    def test_http_server_upload_and_pending_routes(self) -> None:
        with TemporaryDirectory() as tmp:
            service = LucasCloudService(Path(tmp))
            token = service.create_workspace("danny", "Danny")["token"]
            server = make_server(Path(tmp), host="127.0.0.1", port=0)
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            host, port = server.server_address
            try:
                conn = HTTPConnection(host, port)
                body = json.dumps(
                    {
                        "assigned_person": "Danny",
                        "images": [{"name": "front.jpg", "image": "data:image/jpeg;base64,anBn"}],
                    }
                )
                conn.request(
                    "POST",
                    "/api/v1/workspaces/danny/photos/upload",
                    body=body,
                    headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                )
                response = conn.getresponse()
                payload = json.loads(response.read().decode("utf-8"))
                self.assertEqual(response.status, 200)
                self.assertTrue(payload["ok"])

                conn.request("GET", "/api/v1/workspaces/danny/photos/pending", headers={"Authorization": f"Bearer {token}"})
                response = conn.getresponse()
                pending = json.loads(response.read().decode("utf-8"))
                self.assertEqual(response.status, 200)
                self.assertEqual(len(pending["photos"]), 1)
                conn.close()
            finally:
                server.shutdown()
                server.server_close()
                thread.join(timeout=5)


if __name__ == "__main__":
    unittest.main()
