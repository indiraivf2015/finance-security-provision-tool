# Deploying on EC2 (Ubuntu + Nginx)

The tool is one static file (`index.html`). Servers clone this repo and Nginx serves only that file.
No Node, database or backend on the server.

## Current servers
| | UAT | Production |
|---|---|---|
| Server | <UAT server> (<user>) | <production server> (<user>) |
| Folder | <UAT folder> | <production folder> |
| Domain | uatsecprov.indiraivf.in | secprov.indiraivf.in (to be pointed by IT) |
| HTTPS | On Nginx, company certificate /etc/nginx/ssl/indiraivf.in/ | On the load balancer; Nginx listens on port 80 only |
| Config | /etc/nginx/sites-available/secprov | /etc/nginx/sites-available/secprov.indiraivf.in |

## Update the tool (both servers)
```
cd <folder> && git pull
sha256sum index.html      # must equal the hash printed by npm run build
```
No Nginx restart needed (Cache-Control: no-cache).

## New setup
1. `git clone https://github.com/indiraivf2015/finance-security-provision-tool.git` (repo is public; if made private, use a read-only deploy key and `git config core.sshCommand "ssh -i ~/.ssh/secprov_deploy -o IdentitiesOnly=yes"`).
2. Nginx must be able to traverse the folders: `namei -l <folder>/index.html`, then `chmod o+x` on any folder ending in `---`.
3. Copy `uat.nginx.conf` or `prod.nginx.conf` (exactly what is deployed) to /etc/nginx/sites-available/, adjust server_name and root, `ln -s` into sites-enabled.
4. `sudo nginx -t` and only if successful `sudo systemctl reload nginx`.
5. Check: `curl -s -o /dev/null -w "%{http_code}\n" -H "Host: <domain>" http://127.0.0.1/` -> 200 (use 127.0.0.1, not localhost: localhost goes over IPv6).

## Lessons from UAT
- https showing another app = no 443 block for this domain; Nginx falls back to the first 443 site.
- certbot is not needed: the servers use the company indiraivf.in certificate (UAT) or the load balancer (production).
