# Publication

The owner has requested that agents handle Apps Script publication themselves after requested application changes. Use the existing GitHub Actions workflow; Google credentials stay in GitHub, not in the Codex workspace.

A GitHub connector error or HTTP 403 does not establish that the environment's GitHub CLI lacks Actions access. Before asking the owner to run the workflow manually, check `gh auth status` without exposing tokens and try the CLI:

```sh
gh workflow run apps-script.yml --repo kichaad/QueueDesignerAPI --ref master -f operation=publish
```

First validate the change and push the intended commit. Confirm the dispatched run's head SHA matches that commit. Watch the run to completion and report its URL and actual result. Use `operation=verify` for access-only diagnosis without publication. Do not rerun successful publication unnecessarily.

If CLI dispatch also fails, report the specific operation and HTTP error. Git push access and secret-settings access are separate from Actions dispatch access. Never print credentials or request secret values in chat.
