from __future__ import annotations

import uvicorn


def main() -> int:
    uvicorn.run("office_agent.api.app:create_app", host="127.0.0.1", port=8000, factory=True, reload=False)
    return 0
