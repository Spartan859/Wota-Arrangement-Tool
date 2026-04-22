# Compatibility shim — the implementation lives in the wota_tool package.
from wota_tool.models import (  # noqa: F401
    ArrangementState,
    Block,
    HEADERS,
    PURE_ACTION_LYRIC,
    SECTION_MAP,
    TITLE_PATTERN,
)
from wota_tool.app import WotaArrangementTool  # noqa: F401


def main() -> None:
    WotaArrangementTool.configure_stdio()
    tool = WotaArrangementTool()
    tool.run()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n 👋 已退出")
