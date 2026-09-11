"""Subprocess fixture for pipe, deadline and process-ownership tests."""

import json
import os
import sys
import time
from pathlib import Path

from outlook_local_mcp.models import WorkerRequest, WorkerSuccess

for line in sys.stdin:
    request = WorkerRequest.model_validate_json(line)
    arguments = request.arguments
    marker = arguments.get("marker")
    if isinstance(marker, str):
        Path(marker).write_text(str(os.getpid()))
    if len(sys.argv) > 1 and sys.argv[1] == "block":
        Path(sys.argv[2]).write_text(str(os.getpid()))
        time.sleep(60)
    if arguments.get("crash"):
        sys.exit(2)
    delay = arguments.get("delay", 0)
    if isinstance(delay, (int, float)):
        time.sleep(delay)
    response = {"pid": os.getpid(), "available": True, "version": "synthetic", "store_count": 1}
    if len(sys.argv) > 1 and sys.argv[1] == "invalid":
        response = {"invalid": "synthetic-private-response"}
    print(json.dumps(WorkerSuccess(result=response).model_dump(mode="json")), flush=True)
