---
aliases:
  - GitHub
tags:
  - dev
---
# Git

## Daily work

```bash
git status
git add -p                      # stage hunks interactively
git commit -m "Message"
git pull --rebase
git push -u origin feature/x
```

## Branches

```bash
git switch -c feature/x         # create and switch
git branch -d feature/x         # delete merged branch
git rebase main                 # replay branch on main
git merge --no-ff feature/x
```

## Undo

```bash
git restore file.txt            # discard working changes
git restore --staged file.txt   # unstage
git commit --amend --no-edit
git revert <sha>                # new commit that undoes <sha>
git reflog                      # find lost commits
```

## Inspect

```bash
git log --oneline --graph --decorate -20
git diff main...feature/x
git blame -L 10,20 file.ps1
git show <sha>:path/file.txt
```

> [!tip]
> `git config --global --add safe.directory <path>` fixes "dubious ownership" on folders owned by another account.
