"""Run the production MCP boundary with a controllable subprocess instead of COM."""

import asyncio
import sys

from outlook_local_mcp.server import serve
from outlook_local_mcp.supervisor import Supervisor

asyncio.run(serve(Supervisor(command=[sys.executable, "-m", "tests.fake_worker", *sys.argv[1:]])))
