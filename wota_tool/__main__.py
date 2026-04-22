from .app import WotaArrangementTool


def main() -> None:
    WotaArrangementTool.configure_stdio()
    tool = WotaArrangementTool()
    tool.run()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n 👋 已退出")
