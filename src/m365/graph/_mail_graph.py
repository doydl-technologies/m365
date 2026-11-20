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
"""Mailbox policy adapter over the shared, redacted Graph transport."""
from ._graph import GraphAPI, GraphError


class MailboxAPI(GraphAPI):
    def _headers(self, headers):
        result = {k: v for k, v in (headers or {}).items() if k.lower() != "prefer"}
        prefer = [part.strip() for k, v in (headers or {}).items() if k.lower() == "prefer" for part in v.split(",") if part.strip()]
        if self.parent.use_immutable_ids:
            prefer.insert(0, 'IdType="ImmutableId"')
        if prefer:
            result["Prefer"] = ", ".join(dict.fromkeys(prefer))
        return result

    def graph_call(self, verbose=False, resource_url=None, method="GET", batch_requests=None,
                   data=None, json=None, headers=None, stream=False, batch_size=20,
                   content_type="application/json", *, preauthenticated=False, response_metadata=False):
        # Immutable Outlook IDs and body-format preferences apply to Graph, not
        # preauthenticated upload URLs. Batch preferences belong on each member.
        if batch_requests is not None:
            batch_requests = [dict(r, headers=self._headers(r.get("headers"))) for r in batch_requests]
        if not preauthenticated:
            headers = self._headers(headers)
        try:
            return super().graph_call(verbose=verbose, resource_url=resource_url, method=method,
                batch_requests=batch_requests, data=data, json=json, headers=headers, stream=stream,
                batch_size=batch_size, content_type=content_type, preauthenticated=preauthenticated,
                response_metadata=response_metadata)
        except GraphError as exc:
            # Preserve the legacy mailbox JSON-error contract for ordinary calls,
            # without logging response bodies/tokens/URLs. Streams must raise:
            # returning an error dictionary as a byte stream is never safe.
            if stream or preauthenticated or response_metadata:
                raise
            return {"error": str(exc), "status_code": exc.status_code, "code": exc.code}
