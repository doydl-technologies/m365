# -*- coding: utf-8 -*-
#
# DOYDL's Microsoft Graph Integration Toolkit — m365
#
# The `m365` package provides reusable Python clients for Microsoft Graph.
# It supports SharePoint document libraries, OneDrive files and folders,
# and Outlook mailboxes through the `m365.graph` namespace.
#
# Applications can search for files and folders, transfer documents, move
# or rename items within a drive, and work with messages and attachments.
# The toolkit separates these operations from tenant-specific authentication,
# configuration, and business logic.
#
# Access is subject to Microsoft Graph capabilities, tenant policies, and
# the permissions granted to the calling application or signed-in user.
# Installing the package does not grant access to Microsoft 365 resources.
#
# Microsoft, Microsoft Graph, Microsoft 365, Outlook, SharePoint, and OneDrive
# are trademarks of Microsoft Corporation. This project is not an official
# Microsoft product and does not imply affiliation, sponsorship, or
# endorsement. Third-party assets remain subject to their own terms and
# notices; the license below does not grant rights to third-party trademarks
# or logos.
#
# Copyright (c) 2026 by DOYDL Technologies. All rights reserved.
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the “Software”), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#
"""Canonical Graph namespace and intentionally removed compatibility boundaries."""
import ast
from email import message_from_bytes
from importlib import import_module
from importlib.machinery import PathFinder
from importlib.resources import files
import os
from pathlib import Path
import pickle
import subprocess
import sys
import tarfile
import unittest
from zipfile import ZipFile


class NamespaceTests(unittest.TestCase):
    def test_shared_storage_engine_remains_internal(self):
        import m365.graph as graph
        from m365.graph._storage import _StorageClient
        self.assertTrue(issubclass(graph.Drive, _StorageClient))
        for name in ("_StorageClient", "StorageClient"):
            self.assertNotIn(name, graph.__all__)
            self.assertNotIn(name, dir(graph))
        self.assertEqual(graph.Drive.move.__module__, "m365.graph._storage")

    def test_artifacts_include_shared_storage_and_public_guide_without_readme(self):
        if not os.environ.get("M365_DIST"):
            self.skipTest("Build artifacts required")
        directory = Path(os.environ["M365_DIST"])
        wheel, = directory.glob("*.whl")
        sdist, = directory.glob("*.tar.gz")
        retired = {"graph.md", "onedrive.md"}
        with ZipFile(wheel) as archive:
            for module in ("_storage.py", "_storage_auth.py"):
                self.assertIn("m365/graph/" + module, archive.namelist())
            self.assertFalse(any(Path(n).name.lower().startswith("readme") for n in archive.namelist()))
            self.assertTrue(retired.isdisjoint(Path(n).name.lower() for n in archive.namelist()))
            metadata, = [n for n in archive.namelist() if n.endswith(".dist-info/METADATA")]
            self.assertFalse((message_from_bytes(archive.read(metadata)).get_payload() or "").strip())
        with tarfile.open(sdist) as archive:
            self.assertTrue(any(n.endswith("/STORAGE.md") for n in archive.getnames()))
            self.assertFalse(any(Path(n).name.lower().startswith("readme") for n in archive.getnames()))
            self.assertTrue(retired.isdisjoint(Path(n).name.lower() for n in archive.getnames()))

    def test_renamed_clients_have_real_canonical_class_names(self):
        from m365.graph import Drive, Mail
        self.assertEqual(Drive.__name__, "Drive")
        self.assertEqual(Mail.__name__, "Mail")
        for cls in (Drive, Mail):
            self.assertIs(pickle.loads(pickle.dumps(cls)), cls)

    def test_renamed_clients_match_implementation_modules(self):
        from m365.graph import Drive, Mail
        from m365.graph.onedrive import Drive as DirectDrive
        from m365.graph.mailbox import Mail as DirectMail
        self.assertIs(Drive, DirectDrive)
        self.assertIs(Mail, DirectMail)
        self.assertIn("Mail", dir(import_module("m365.graph.mailbox")))
        self.assertEqual(import_module("m365.graph.onedrive").__all__, ["Drive"])

    def test_previous_client_names_are_not_compatibility_aliases(self):
        import m365.graph as graph
        for name, suffix in (("OneDrive", "onedrive"), ("Mailbox", "mailbox")):
            with self.subTest(previous=name):
                self.assertFalse(hasattr(graph, name))
                self.assertFalse(hasattr(import_module("m365.graph." + suffix), name))
                with self.assertRaises(ImportError):
                    exec("from m365.graph import " + name, {})

    def test_graph_public_exports_are_only_clients_and_error(self):
        import m365.graph as graph
        self.assertEqual(graph.__all__, ["Site", "Drive", "Mail", "GraphError"])
        self.assertEqual({name for name in dir(graph) if not name.startswith("_")}, set(graph.__all__))

    def test_wildcard_import_does_not_expose_helpers(self):
        scope = {}
        exec("from m365.graph import *", scope)
        self.assertEqual(set(scope) - {"__builtins__"}, {"Site", "Drive", "Mail", "GraphError"})

    def test_old_helper_packages_are_absent(self):
        import importlib.util
        import m365.graph as graph
        for name in ("dateplumb", "syslib", "preprocessing"):
            with self.subTest(helper=name):
                self.assertFalse(hasattr(graph, name))
                self.assertIsNone(importlib.util.find_spec("m365.graph." + name))
                with self.assertRaises(ModuleNotFoundError):
                    import_module("m365.graph." + name)
                with self.assertRaises(ImportError):
                    exec("from m365.graph import " + name, {})

    def test_private_helpers_are_usable_without_public_aliases(self):
        import m365.graph as graph
        for name in ("dateplumb", "syslib", "preprocessing"):
            module = import_module("m365.graph._" + name)
            self.assertTrue(module.__name__.startswith("m365.graph._"))
            self.assertFalse(hasattr(graph, name))
        dates = import_module("m365.graph._dateplumb")
        self.assertEqual(dates.parse("2026-11-19").year, 2026)

    def test_artifacts_ship_private_helpers_and_resources_only(self):
        if not os.environ.get("M365_DIST"):
            self.skipTest("Build artifacts required")
        directory = Path(os.environ["M365_DIST"])
        wheel, = directory.glob("*.whl")
        sdist, = directory.glob("*.tar.gz")
        with ZipFile(wheel) as archive:
            names = archive.namelist()
            for helper in ("dateplumb", "syslib", "preprocessing"):
                self.assertIn("m365/graph/_" + helper + "/__init__.py", names)
                self.assertFalse(any(n.startswith("m365/graph/" + helper + "/") for n in names))
            self.assertTrue(archive.read("m365/graph/_syslib/ref_data/extensions.txt"))
        with tarfile.open(sdist) as archive:
            names = archive.getnames()
            for helper in ("dateplumb", "syslib", "preprocessing"):
                self.assertTrue(any(n.endswith("/src/m365/graph/_" + helper + "/__init__.py") for n in names))
                self.assertFalse(any("/src/m365/graph/" + helper + "/" in n for n in names))

    def child(self, source):
        import m365
        selected = str(Path(m365.__file__).resolve().parents[1])
        source = "import sys; sys.path.insert(0, {!r})\n".format(selected) + source
        result = subprocess.run([sys.executable, "-B", "-c", source],
                                capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_root_import_is_lazy(self):
        self.child("import sys, m365; assert 'm365.graph' not in sys.modules; "
                   "assert 'comoutlook' not in sys.modules; assert 'pythoncom' not in sys.modules")

    def test_public_classes_live_in_canonical_modules(self):
        from m365.graph import Site, Drive, Mail, GraphError
        self.assertEqual(Site.__module__, "m365.graph.sharepoint")
        self.assertEqual(Drive.__module__, "m365.graph.onedrive")
        self.assertEqual(Mail.__module__, "m365.graph.mailbox")
        self.assertEqual(GraphError.__module__, "m365.graph._graph")

    def test_nested_modules_load_once(self):
        for suffix in ("operations.doc_sync", "auth_types", "_syslib._utils", "_dateplumb._advanced"):
            with self.subTest(suffix=suffix):
                a = import_module("m365.graph." + suffix)
                b = import_module("m365.graph." + suffix)
                self.assertIs(a, b)
                self.assertEqual(a.__spec__.name, "m365.graph." + suffix)

    def test_resources_available(self):
        content = files("m365.graph._syslib.ref_data").joinpath("extensions.txt").read_bytes()
        self.assertIn(b".txt", content)

    def test_no_legacy_import_hook_installed(self):
        self.child("import sys, m365.graph; "
                   "assert not any(getattr(f,'_m365_legacy_graph',False) for f in sys.meta_path); "
                   "assert not any(n=='msftgraph' or n.startswith('msftgraph.') for n in sys.modules)")

    def test_old_package_not_in_selected_distribution(self):
        import m365
        selected = str(Path(m365.__file__).resolve().parents[1])
        self.assertIsNone(PathFinder.find_spec("msftgraph", [selected]))

    def test_private_compatibility_helper_removed(self):
        import importlib.util
        self.assertIsNone(importlib.util.find_spec("m365._compat"))

    def test_pickle_uses_canonical_module(self):
        from m365.graph.operations.doc_sync import Upload
        encoded = pickle.dumps(Upload)
        self.assertNotIn(b"msftgraph", encoded)
        self.assertIn(b"m365.graph.operations.doc_sync", encoded)
        self.assertIs(pickle.loads(encoded), Upload)

    def test_graph_does_not_load_outlook(self):
        self.child("import sys; from m365.graph import Site; "
                   "assert 'comoutlook' not in sys.modules; assert 'pythoncom' not in sys.modules")

    def test_star_import_does_not_load_outlook(self):
        self.child("from m365 import *; import sys; assert 'comoutlook' not in sys.modules")

    def test_removed_com_namespace_cannot_be_imported(self):
        import importlib.util
        self.assertIsNone(importlib.util.find_spec("m365.outlook_com"))

    def test_unknown_root_attribute_raises(self):
        import m365
        with self.assertRaises(AttributeError):
            m365.not_a_backend

    def test_backend_has_no_old_absolute_imports(self):
        root = Path(__file__).resolve().parents[2] / "src" / "m365" / "graph"
        self.assertTrue(root.is_dir())
        for file in root.rglob("*.py"):
            for node in ast.walk(ast.parse(file.read_text(encoding="utf-8-sig"))):
                if isinstance(node, ast.ImportFrom) and node.level == 0:
                    self.assertFalse((node.module or "").startswith("msftgraph"), str(file))
                if isinstance(node, ast.Import):
                    self.assertFalse(any(n.name.startswith("msftgraph") for n in node.names), str(file))

    def test_wheel_and_sdist_have_no_legacy_or_com_backend(self):
        if not os.environ.get("M365_DIST"):
            self.skipTest("Build artifacts required")
        directory = Path(os.environ["M365_DIST"])
        wheel, = directory.glob("*.whl")
        sdist, = directory.glob("*.tar.gz")
        def check(names):
            self.assertFalse(any(n.startswith("msftgraph/") or "/msftgraph/" in n or
                                 "m365/_compat.py" in n or "comoutlook/" in n or
                                 "m365/outlook_com" in n for n in names))
            self.assertFalse(any("/tests/com/" in n for n in names))
        with ZipFile(wheel) as archive:
            names = archive.namelist()
            check(names)
            self.assertIn("m365/graph/__init__.py", names)
            self.assertIn("m365/graph/sharepoint.py", names)
            metadata, = [n for n in names if n.endswith(".dist-info/METADATA")]
            info = message_from_bytes(archive.read(metadata))
        self.assertEqual(info["Name"], "m365")
        self.assertEqual(info["Version"], "2.9.4")
        self.assertNotIn("outlook-com", info.get_all("Provides-Extra") or [])
        auth, = [r for r in info.get_all("Requires-Dist") if r.startswith("azauthlib")]
        self.assertNotIn(";", auth)
        with tarfile.open(sdist, "r:gz") as archive:
            names = archive.getnames()
            check(names)
            self.assertTrue(any(n.endswith("/src/m365/graph/sharepoint.py") for n in names))
            self.assertFalse(any(n.endswith("/m365/graph/sharepoint.py") and "/src/" not in n
                                 for n in names))

    def test_root_exposes_only_graph_and_matches_release_version(self):
        import m365
        self.assertEqual(m365.__version__, "2.9.4")
        self.assertEqual(m365.graph.__version__, m365.__version__)
        self.assertEqual(m365.__all__, ["graph"])
        with self.assertRaises(AttributeError):
            m365.outlook_com

    def test_src_is_the_only_implementation_tree(self):
        root = Path(__file__).resolve().parents[2]
        self.assertTrue((root / "src/m365/graph/sharepoint.py").is_file())
        self.assertFalse((root / "m365").exists())
        self.assertFalse((root / "src/m365/outlook_com").exists())

    def test_installed_wheel_has_canonical_package_not_src_prefix(self):
        if not os.environ.get("M365_DIST"):
            self.skipTest("Build artifacts required")
        wheel, = Path(os.environ["M365_DIST"]).glob("*.whl")
        with ZipFile(wheel) as archive:
            self.assertIn("m365/__init__.py", archive.namelist())
            self.assertFalse(any(n.startswith("src/") for n in archive.namelist()))

    def test_metadata_declares_no_direct_com_dependencies(self):
        if not os.environ.get("M365_DIST"):
            self.skipTest("Build artifacts required")
        wheel, = Path(os.environ["M365_DIST"]).glob("*.whl")
        with ZipFile(wheel) as archive:
            path, = [n for n in archive.namelist() if n.endswith(".dist-info/METADATA")]
            metadata = message_from_bytes(archive.read(path))
        self.assertFalse(any(r.startswith(("pywin32", "python-dateutil"))
                             for r in metadata.get_all("Requires-Dist")))

    def test_readmes_and_long_description_are_not_distributed(self):
        if not os.environ.get("M365_DIST"):
            self.skipTest("Build artifacts required")
        directory = Path(os.environ["M365_DIST"])
        wheel, = directory.glob("*.whl")
        sdist, = directory.glob("*.tar.gz")
        with ZipFile(wheel) as archive:
            self.assertFalse(any(Path(n).name.startswith("README") for n in archive.namelist()))
            path, = [n for n in archive.namelist() if n.endswith(".dist-info/METADATA")]
            info = message_from_bytes(archive.read(path))
            self.assertFalse(info.get_payload(decode=True))
        with tarfile.open(sdist, "r:gz") as archive:
            self.assertFalse(any(Path(n).name.startswith("README") for n in archive.getnames()))
            path, = [n for n in archive.getnames() if n.count("/") == 1 and n.endswith("/PKG-INFO")]
            info = message_from_bytes(archive.extractfile(path).read())
            self.assertFalse(info.get_payload(decode=True))
