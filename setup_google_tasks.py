import argparse
import os
import sys
from pathlib import Path

WORKSPACE_DIR = Path(__file__).resolve().parent
VENV_SITE_PACKAGES = WORKSPACE_DIR / ".venv" / "Lib" / "site-packages"
BUNDLED_SITE_PACKAGES = Path(
    r"C:\Users\laput\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\Lib\site-packages"
)

for site_packages in (VENV_SITE_PACKAGES, BUNDLED_SITE_PACKAGES):
    if site_packages.exists():
        sys.path.insert(0, str(site_packages))

from agent_config import apply_secret_defaults, load_agent_env
import summarize_note5 as summarize_note


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create or refresh the Google Tasks OAuth token."
    )
    parser.add_argument(
        "--credentials",
        default=os.getenv("GOOGLE_TASKS_CREDENTIALS", "credentials.json"),
        help="OAuth desktop client JSON downloaded from Google Cloud.",
    )
    parser.add_argument(
        "--token",
        default=os.getenv("GOOGLE_TASKS_TOKEN", "token_google_tasks.json"),
        help="OAuth token JSON to create or refresh.",
    )
    parser.add_argument(
        "--tasklist",
        default=os.getenv("GOOGLE_TASKS_LIST_ID", "@default"),
        help="Google Tasks tasklist id. Use @default for the default list.",
    )
    parser.add_argument(
        "--add-test-task",
        action="store_true",
        help="Add one test task after authentication.",
    )
    return parser.parse_args()


def main() -> None:
    load_agent_env()
    apply_secret_defaults()
    args = parse_args()
    credentials_path = Path(args.credentials)
    token_path = Path(args.token)

    if not credentials_path.exists():
        print(f"[missing] {credentials_path}")
        print("Download an OAuth desktop client JSON from Google Cloud.")
        print("Save it in the external secret directory shown by: .\agent.ps1 secrets-status")
        raise SystemExit(2)

    service = summarize_note.build_google_tasks_service(credentials_path, token_path, interactive_auth=True)
    tasklists = service.tasklists().list(maxResults=20).execute().get("items", [])

    print("[ok] Google Tasks authentication succeeded.")
    print(f"token: {token_path}")
    print("tasklists:")
    for tasklist in tasklists:
        print(f"- {tasklist.get('title')} ({tasklist.get('id')})")

    if args.add_test_task:
        created = (
            service.tasks()
            .insert(tasklist=args.tasklist, body={"title": "Codex Google Tasks connection test"})
            .execute()
        )
        print(f"[ok] Added test task: {created.get('title')} ({created.get('id')})")


if __name__ == "__main__":
    main()
