# Smart PM — Online Deployment

This build is prepared to run on an online server instead of being tied to localhost.

## What is already changed
- Uses `PORT` from the hosting platform (default 3050).
- Listens on `0.0.0.0` so the service can receive external traffic.
- Uses `PM_DATA_ROOT` for persistent PM data, Excel imports, users/config, branding and history.
- Includes `/health` for hosting health checks.
- Includes proxy/HTTPS-aware security headers.
- Includes Dockerfile and docker-compose.yml.

## Important: persistent data
Do not deploy without persistent storage. Mount a persistent disk/volume to `/app/data` and set:

`PM_DATA_ROOT=/app/data`

Otherwise imported Excel plans and execution history can be lost when a cloud instance is replaced.

## Option A — Windows/Linux VPS with Docker
1. Install Docker.
2. Copy this application folder to the server.
3. In the folder run: `docker compose up -d --build`
4. Allow TCP 3050 temporarily in the firewall for testing.
5. Open `http://SERVER_PUBLIC_IP:3050`.
6. For production, connect a domain such as `pm.company.com` through an HTTPS reverse proxy and close direct public access to port 3050.

The named Docker volume `smart_pm_data` keeps the PM data between container rebuilds.

## Option B — Cloud container host
Deploy using the included Dockerfile. Configure:
- `PORT`: normally supplied by the host.
- `HOST=0.0.0.0`
- `PM_DATA_ROOT=/app/data` (or the mount path supplied by the host)
- `PUBLIC_URL=https://your-domain.example`

Attach a persistent disk/volume at the same path used by `PM_DATA_ROOT`.
Set the health-check path to `/health`.

## Existing local data
If you want the online server to start with your current PM plans/history, copy the current application's `data` folder into the persistent server volume before first production use. Do not overwrite live online data later without a backup.

## Production checklist
- Change all default/test passwords before exposing the site publicly.
- Use HTTPS.
- Keep regular backups of the persistent data volume.
- Restrict server/firewall administration access.
- Test Admin, Technician and Viewer permissions after deployment.
- Test Excel import, Done/Deferred, reports, annual plan, calendar and branding before users start production work.

## Note about login sessions
Current login sessions are held in server memory. Users remain authenticated while the Node process is running, but a server restart will require them to log in again. PM plans/history are persistent because those are stored under `PM_DATA_ROOT`.
