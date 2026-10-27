# Deploying on EC2 (Ubuntu + Nginx)

The site is a few static files at the top of this repo: `index.html` (start page), `security.html`, and one
`<provision>.html` per provision. Servers clone this repo; Nginx serves only `/` and those top-level `.html` files.
No Node, database or backend on the server.

## Current servers
| | UAT | Production |
|---|---|---|
| Server | <UAT server> (<user>) | <production server> (<user>) |
| Folder | <UAT folder> | <production folder> |
| Domain | uatsecprov.indiraivf.in | secprov.indiraivf.in |
| HTTPS | On Nginx, company certificate /etc/nginx/ssl/indiraivf.in/ | On the AWS load balancer; Nginx listens on port 80 only (no redirect: it would loop) |
| Config | /etc/nginx/sites-available/secprov | /etc/nginx/sites-available/secprov.indiraivf.in |
| Repo access | read-only deploy key ~/.ssh/secprov_deploy | read-only deploy key ~/.ssh/secprov_deploy |

## Update the site (both servers; UAT first, finance checks, then production)
```
cd <folder>
git status                 # must say "nothing to commit"; a hand-edited file blocks the pull
git pull
sha256sum -c SHA256SUMS    # every page must say OK
```
No Nginx restart needed (Cache-Control: no-cache). If a release also changes the Nginx rules, change Nginx FIRST
(see below), then `git pull`.

## Changing the Nginx rules
```
sudo cp <folder>/deploy-ec2/<uat|prod>.nginx.conf /tmp/new.conf   # then set server_name and root to this server's values
sudo cp /etc/nginx/sites-available/<config> ~/<config>.bak        # backup
sudo nano /etc/nginx/sites-available/<config>                     # apply the change
sudo nginx -t && sudo systemctl reload nginx                      # reload ONLY if the test passes
```
If `nginx -t` fails: restore the backup and do not reload.

## Checks after any change (on the server; use 127.0.0.1, not localhost: localhost goes over IPv6 to another site)
```
for u in / /index.html /security.html /.git/config /src/security/ui.html /SHA256SUMS /dist/security.html; do
  printf '%-24s ' "$u"; curl -s -o /dev/null -w "%{http_code}\n" -H "Host: <domain>" http://127.0.0.1$u
done
```
Expected: `/`, `/index.html` and every page in SHA256SUMS -> 200; everything else -> 404.
UAT redirects port 80 to https, so on UAT use `https://127.0.0.1$u -k` (or `--resolve <domain>:443:127.0.0.1`).

## Repo access with a deploy key (needed because the repo is private)
On the server, as the deploy user:
```
ssh-keygen -t ed25519 -f ~/.ssh/secprov_deploy -N "" -C "secprov-<server>-readonly"
chmod 600 ~/.ssh/secprov_deploy
cat ~/.ssh/secprov_deploy.pub     # GitHub repo > Settings > Deploy keys > Add; "Allow write access" UNticked
ssh -i ~/.ssh/secprov_deploy -o IdentitiesOnly=yes -T git@github.com
# first connect: accept only if the ED25519 fingerprint is SHA256:+DiY3wvvV6TuJJhbpZisF/zLDA0zPMSvHdkr4UvCOqU
cd <folder>
git remote set-url origin git@github.com:indiraivf2015/finance-security-provision-tool.git
git config core.sshCommand "ssh -i ~/.ssh/secprov_deploy -o IdentitiesOnly=yes"
git fetch                          # proves access without changing the served files
```
Each server needs its own key (GitHub does not allow the same deploy key twice).

## New setup
1. Clone with the deploy key (above): `git clone git@github.com:indiraivf2015/finance-security-provision-tool.git`.
2. Nginx must be able to traverse the folders: `namei -l <folder>/index.html`, then `chmod o+x` on any folder ending in `---`.
3. Copy `uat.nginx.conf` or `prod.nginx.conf` to /etc/nginx/sites-available/, set server_name and root, `ln -s` into sites-enabled.
4. `sudo nginx -t` and only if successful `sudo systemctl reload nginx`.
5. Run the checks above.

## Lessons
- https showing another app = no 443 block for this domain; Nginx falls back to the first 443 site.
- certbot is not needed: the servers use the company indiraivf.in certificate (UAT) or the load balancer (production).
- Production: never add an http->https redirect in Nginx; the load balancer already does HTTPS.
