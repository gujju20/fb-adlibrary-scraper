"""Ad Library Studio web server. Run: python web_app.py"""
import base64
from contextlib import asynccontextmanager
import hmac
import json
import os
from pathlib import Path
import re
import secrets
from urllib.parse import parse_qs, urlsplit

from starlette.applications import Starlette
from starlette.middleware import Middleware
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import FileResponse, JSONResponse, PlainTextResponse
from starlette.routing import Route, Mount
from starlette.staticfiles import StaticFiles

from jobs import JobManager, BUSY
from meta_ads_export import source_settings

ROOT = Path(__file__).resolve().parent


class SessionBoundary(BaseHTTPMiddleware):
    async def dispatch(self, request, call_next):
        if request.url.path == '/health':
            return JSONResponse({'status': 'ok'})
        password = os.environ.get('APP_PASSWORD', '')
        if password:
            expected = 'Basic ' + base64.b64encode(('admin:' + password).encode()).decode()
            if not hmac.compare_digest(request.headers.get('authorization', '').encode(), expected.encode()):
                return PlainTextResponse('Sign in with username admin and your app password.', 401,
                                         headers={'WWW-Authenticate': 'Basic realm="Ad Library Studio"'})
        if request.method not in ('GET', 'HEAD'):
            origin = request.headers.get('origin')
            if origin and origin.rstrip('/') != str(request.base_url).rstrip('/'):
                return JSONResponse({'error': 'Cross-site requests are not allowed'}, 403)
            if request.headers.get('content-type', '').split(';')[0] != 'application/json':
                return JSONResponse({'error': 'Expected application/json'}, 415)
            try:
                content_length = int(request.headers.get('content-length', '0') or '0')
            except ValueError:
                return JSONResponse({'error': 'Invalid content length'}, 400)
            if content_length > 16384:
                return JSONResponse({'error': 'Request too large'}, 413)
        session = request.cookies.get('studio_session', '')
        new_session = not re.fullmatch(r'[a-f0-9]{64}', session)
        if new_session:
            session = secrets.token_hex(32)
        request.state.owner = session
        response = await call_next(request)
        if new_session:
            response.set_cookie('studio_session', session, max_age=7 * 86400, httponly=True,
                                secure=request.url.scheme == 'https', samesite='strict')
        response.headers['X-Content-Type-Options'] = 'nosniff'
        response.headers['Referrer-Policy'] = 'same-origin'
        response.headers['Content-Security-Policy'] = "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'self'; form-action 'self'"
        response.headers['Cache-Control'] = 'no-store'
        return response


def manager(request):
    return request.app.state.jobs


def index(request):
    return FileResponse(ROOT / 'web' / 'index.html')


def list_jobs(request):
    return JSONResponse([{key: job[key] for key in ('id', 'country', 'created', 'scrape', 'assets')}
                         for job in manager(request).list(request.state.owner)])


async def create_job(request):
    try:
        raw = bytearray()
        async for chunk in request.stream():
            raw.extend(chunk)
            if len(raw) > 16384:
                return JSONResponse({'error': 'Request too large'}, 413)
        data = json.loads(raw)
        if not isinstance(data, dict):
            raise ValueError('Expected a JSON object')
        url = data.get('url', '')
        if not isinstance(url, str) or len(url) > 8192:
            raise ValueError('Enter a valid Facebook Ad Library URL')
        url, code, country = source_settings(url.strip())
        params = parse_qs(urlsplit(url).query)
        # Keep the original exporter's active-ad contract explicit.
        if params.get('active_status', ['active'])[0] not in ('active', 'all'):
            raise ValueError('This exporter reads active ads. Use an active or all-status search link.')
        limit = data.get('limit', 0)
        if type(limit) is not int or limit < 0 or limit > 100000:
            raise ValueError('Ad limit must be a whole number from 0 to 100000; 0 means all results')
        key = manager(request).create(request.state.owner, url, code, country, limit)
        return JSONResponse(manager(request).get(key, request.state.owner), 202)
    except (ValueError, TypeError) as exc:
        return JSONResponse({'error': str(exc)}, 400)


def get_job(request):
    try:
        return JSONResponse(manager(request).get(request.path_params['key'], request.state.owner))
    except KeyError:
        return JSONResponse({'error': 'Export not found or session expired'}, 404)


def start_assets(request):
    try:
        manager(request).start_assets(request.path_params['key'], request.state.owner)
        return get_job(request)
    except KeyError:
        return JSONResponse({'error': 'Export not found'}, 404)
    except ValueError as exc:
        return JSONResponse({'error': str(exc)}, 409)


def cancel_job(request):
    try:
        manager(request).cancel(request.path_params['key'], request.state.owner)
        return get_job(request)
    except KeyError:
        return JSONResponse({'error': 'Export not found'}, 404)


def download_file(request):
    try:
        job = manager(request).get(request.path_params['key'], request.state.owner)
    except KeyError:
        return JSONResponse({'error': 'Export not found'}, 404)
    kind = request.path_params['kind']
    if kind == 'logs':
        return PlainTextResponse('\n'.join(item['text'] for item in job['logs']),
                                 headers={'Content-Disposition': 'attachment; filename="activity.txt"'})
    if kind not in ('xlsx', 'zip') or not job[kind + '_ready']:
        return JSONResponse({'error': 'File is not ready'}, 409)
    name = 'ads.xlsx' if kind == 'xlsx' else 'assets.zip'
    mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if kind == 'xlsx' else 'application/zip'
    return FileResponse(manager(request).directory(job['id']) / name, media_type=mime,
                        filename=f"{job['code'].lower()}_{job['id'][:8]}_{name}")


def create_app(data_dir=None):
    @asynccontextmanager
    async def lifespan(app):
        app.state.jobs = JobManager(data_dir or os.getenv('DATA_DIR', ROOT / 'data'),
                                    workers=int(os.getenv('JOB_CONCURRENCY', '1')))
        yield
        app.state.jobs.close()
    return Starlette(lifespan=lifespan, middleware=[Middleware(SessionBoundary)], routes=[
        Route('/', index), Route('/health', index),
        Route('/api/jobs', list_jobs, methods=['GET']),
        Route('/api/jobs', create_job, methods=['POST']),
        Route('/api/jobs/{key}', get_job),
        Route('/api/jobs/{key}/assets', start_assets, methods=['POST']),
        Route('/api/jobs/{key}/cancel', cancel_job, methods=['POST']),
        Route('/api/jobs/{key}/download/{kind}', download_file),
        Mount('/static', app=StaticFiles(directory=ROOT / 'web')),
    ])


app = create_app()

if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host=os.getenv('HOST', '127.0.0.1'), port=int(os.getenv('PORT', '8000')),
                proxy_headers=True, timeout_graceful_shutdown=55)
