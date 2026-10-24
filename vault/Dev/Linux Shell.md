---
aliases:
  - bash
  - ssh
  - nginx
  - systemctl
  - apt
tags:
  - dev
  - shell
---
# Linux Shell

The aliases make this note appear whenever a window title or copied text contains *ssh*, *bash*, *nginx*, *systemctl* or *apt*. That is the clipboard mode in action: copy a command inside an SSH session and the note is here.

## Services (systemd)

```bash
sudo systemctl status nginx
sudo systemctl restart nginx
sudo systemctl enable --now nginx
sudo journalctl -u nginx -n 100 --no-pager
sudo journalctl -f
```

## nginx

```bash
sudo nginx -t                      # test configuration
sudo nginx -s reload
sudo tail -f /var/log/nginx/error.log
ls /etc/nginx/sites-enabled/
```

## Packages (Debian, Ubuntu)

```bash
sudo apt update && sudo apt upgrade -y
apt list --upgradable
sudo apt install htop
dpkg -l | grep nginx
```

## Files and processes

```bash
du -sh * | sort -h
find /var/log -name "*.log" -mtime +7
grep -rn "error" /var/log/nginx/
ps aux --sort=-%mem | head
ss -tulpn
```

## ssh

```bash
ssh -i ~/.ssh/id_ed25519 user@host
ssh -L 8080:localhost:80 user@host    # local port forward
scp file.txt user@host:/tmp/
ssh-copy-id user@host
```

Related: [[PowerShell]], [[Docker]].
