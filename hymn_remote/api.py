"""FastAPI application for mobile remote control."""
import os
from pathlib import Path

from fastapi import FastAPI, Header, HTTPException, Request, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from typing import Any

from pydantic import BaseModel

from .resolver import list_books, list_entries_for_book
from .state import RemoteControlState


class OpenRequest(BaseModel):
    book: str
    num: str = ''


class SetlistOpenRequest(BaseModel):
    index: int


class DisplayControlRequest(BaseModel):
    action: str  # close | black


class LyricControlRequest(BaseModel):
    action: str
    value: Any = None


class PrefsRequest(BaseModel):
    overlay_auto_unblack: Any = None


def _static_dir():
    return Path(__file__).resolve().parent / 'static'


def _no_store_headers():
    return {
        'Cache-Control': 'no-store, no-cache, must-revalidate, max-age=0',
        'Pragma': 'no-cache',
        'Expires': '0',
    }


def _html_page(path: Path):
    return FileResponse(
        path,
        media_type='text/html; charset=utf-8',
        headers=_no_store_headers(),
    )


def _public_prefix(request: Request) -> str:
    root = str(request.scope.get('root_path') or '').rstrip('/')
    if root:
        return root
    referer = request.headers.get('referer') or ''
    try:
        from urllib.parse import urlparse
        path = urlparse(referer).path or ''
    except Exception:
        path = ''
    for marker in ('/admin', '/view', '/api/'):
        if marker in path:
            return path.split(marker)[0].rstrip('/')
    return ''


def _viewer_text_url(request: Request) -> str:
    try:
        from version import APP_VERSION
        ver = int(APP_VERSION or 0)
    except Exception:
        ver = 0
    prefix = _public_prefix(request)
    return f'{prefix}/view/text?v={ver}'


def _client_ip(request: Request) -> str:
    if request.client:
        return request.client.host or ''
    return ''


def create_app(state: RemoteControlState, path_prefix: str = '') -> FastAPI:
    app = FastAPI(title='Hymn Search Remote', docs_url='/api/docs', redoc_url=None)
    static = _static_dir()

    def _check_token(x_api_token: str = Header(default='')):
        if not state.check_token(x_api_token):
            raise HTTPException(status_code=401, detail='Invalid API token')

    @app.get('/api/health')
    def health():
        try:
            from version import APP_VERSION, BUILD_DATE
        except ImportError:
            APP_VERSION, BUILD_DATE = 0, ''
        snap = state.snapshot()
        return {
            'ok': True,
            'app_version': APP_VERSION,
            'build_date': BUILD_DATE,
            **snap,
        }

    @app.get('/api/books')
    def books(x_api_token: str = Header(default='')):
        _check_token(x_api_token)
        return list_books(state.get_books())

    @app.get('/api/entries')
    def entries(book: str = '', q: str = '', x_api_token: str = Header(default='')):
        _check_token(x_api_token)
        return list_entries_for_book(
            state.get_books(), book, q, num_mode=state.get_hymn_num_mode(),
        )

    @app.get('/api/preview')
    def lyric_preview(rev: int = -1, view: int = -1, x_api_token: str = Header(default='')):
        _check_token(x_api_token)
        return state.get_lyric_preview(
            since_rev=rev if rev >= 0 else None,
            since_view=view if view >= 0 else None,
        )

    @app.post('/api/preview/control')
    def lyric_control(
        body: LyricControlRequest,
        request: Request,
        x_api_token: str = Header(default=''),
    ):
        _check_token(x_api_token)
        if not state.is_accepting():
            return JSONResponse(
                status_code=503,
                content={
                    'status': 'not_accepting',
                    'message': '桌面端暫停接受請求，請聯絡操作員',
                },
            )
        return state.dispatch_lyric_control(body.action, body.value, _client_ip(request))

    @app.post('/api/prefs')
    def set_prefs(
        body: PrefsRequest,
        request: Request,
        x_api_token: str = Header(default=''),
    ):
        _check_token(x_api_token)
        if body.overlay_auto_unblack is not None:
            state.set_overlay_auto_unblack(bool(body.overlay_auto_unblack), persist=True)
        return {'ok': True, **state.snapshot()}

    @app.websocket('/api/preview/ws')
    async def lyric_ws(websocket: WebSocket, token: str = ''):
        hdr = websocket.headers.get('x-api-token', '')
        if not state.check_token(token or hdr):
            await websocket.close(code=4401)
            return
        await websocket.accept()
        client = websocket.client.host if websocket.client else ''
        state.register_preview_ws(websocket)
        try:
            snap = state.get_lyric_preview()
            await websocket.send_json({'type': 'preview', **snap})
            while True:
                data = await websocket.receive_json()
                if not isinstance(data, dict):
                    continue
                action = str(data.get('action') or '').strip()
                if not action:
                    continue
                state.dispatch_lyric_control(action, data.get('value'), client)
        except WebSocketDisconnect:
            return
        except Exception:
            try:
                await websocket.close()
            except Exception:
                pass
        finally:
            state.unregister_preview_ws(websocket)

    @app.post('/api/open')
    def open_item(body: OpenRequest, request: Request, x_api_token: str = Header(default='')):
        _check_token(x_api_token)
        if not state.is_accepting():
            return JSONResponse(
                status_code=503,
                content={
                    'status': 'not_accepting',
                    'message': '桌面端暫停接受請求，請聯絡操作員',
                },
            )
        result = state.dispatch_open_request(body.book, body.num, _client_ip(request))
        if hasattr(state, 'touch_preview_client'):
            state.touch_preview_client()
        return result

    @app.post('/api/next')
    def next_item(request: Request, x_api_token: str = Header(default='')):
        _check_token(x_api_token)
        if not state.is_accepting():
            return JSONResponse(
                status_code=503,
                content={
                    'status': 'not_accepting',
                    'message': '桌面端暫停接受請求，請聯絡操作員',
                },
            )
        return state.dispatch_next(_client_ip(request))

    @app.get('/api/setlist')
    def setlist(x_api_token: str = Header(default='')):
        _check_token(x_api_token)
        return state.get_setlist_info()

    @app.post('/api/setlist/open')
    def setlist_open(body: SetlistOpenRequest, request: Request, x_api_token: str = Header(default='')):
        _check_token(x_api_token)
        if not state.is_accepting():
            return JSONResponse(
                status_code=503,
                content={
                    'status': 'not_accepting',
                    'message': '桌面端暫停接受請求，請聯絡操作員',
                },
            )
        return state.dispatch_setlist_open(body.index, _client_ip(request))

    @app.post('/api/setlist/add')
    def setlist_add(body: OpenRequest, request: Request, x_api_token: str = Header(default='')):
        _check_token(x_api_token)
        return state.dispatch_setlist_add(body.book, body.num, _client_ip(request))

    @app.post('/api/display')
    def display_control(
        body: DisplayControlRequest,
        request: Request,
        x_api_token: str = Header(default=''),
    ):
        _check_token(x_api_token)
        if not state.is_accepting():
            return JSONResponse(
                status_code=503,
                content={
                    'status': 'not_accepting',
                    'message': '桌面端暫停接受請求，請聯絡操作員',
                },
            )
        return state.dispatch_display_control(body.action, _client_ip(request))

    @app.get('/api/log')
    def get_log(limit: int = 50, x_api_token: str = Header(default='')):
        _check_token(x_api_token)
        return state.operation_log.list_entries(limit=min(limit, 200))

    @app.get('/api/pending')
    def pending(x_api_token: str = Header(default='')):
        _check_token(x_api_token)
        return [
            {
                'id': r.id,
                'book_ref': r.book_ref,
                'num_ref': r.num_ref,
                'summary': r.summary,
                'match_count': len(r.matches),
            }
            for r in state.list_pending()
        ]

    @app.post('/api/pending/{req_id}/approve')
    def approve(req_id: str, request: Request, x_api_token: str = Header(default='')):
        _check_token(x_api_token)
        return state.approve_pending(req_id, _client_ip(request))

    @app.post('/api/pending/{req_id}/reject')
    def reject(req_id: str, request: Request, x_api_token: str = Header(default='')):
        _check_token(x_api_token)
        return state.reject_pending(req_id, _client_ip(request))

    @app.get('/api/view/now')
    def view_now(request: Request):
        if not state.api_enabled:
            return JSONResponse(
                status_code=503,
                content={'status': 'disabled', 'open': False, 'text': '', 'web_url': '', 'show_debug': False},
                headers=_no_store_headers(),
            )
        snap = state.get_lyric_preview()
        heading = str(snap.get('heading') or snap.get('name') or '').strip()
        text = str(snap.get('text') or '').strip()
        has_lyrics = bool(text or heading)
        data = {
            'open': bool(snap.get('open') or has_lyrics),
            'heading': heading,
            'title': heading,
            'label': heading,
            'name': snap.get('name') or '',
            'book': snap.get('book') or '',
            'text': text,
            'bg': snap.get('bg') or '#050505',
            'bg_image': bool(snap.get('bg_image')),
            'bg_image_rev': snap.get('bg_image_rev') or 0,
            'text_color': snap.get('text_color') or '#ffffff',
            'align': snap.get('align') or 'left',
            'letter_spacing': snap.get('letter_spacing') or 100,
            'line_height': snap.get('line_height') or 135,
            'margins': snap.get('margins') or {
                'top': 56, 'right': 56, 'bottom': 56, 'left': 56,
            },
            'margin': snap.get('margin') if snap.get('margin') is not None else 56,
            'font_family': snap.get('font_family') or 'Microsoft YaHei UI',
            'blacked': bool(snap.get('blacked')),
            'rev': snap.get('rev'),
            'view_rev': snap.get('view_rev'),
            'mapped': has_lyrics,
            'viewer_source': 'text',
            'web_url': _viewer_text_url(request) if has_lyrics else '',
            'show_debug': state.get_viewer_show_debug(),
        }
        return JSONResponse(data, headers=_no_store_headers())

    @app.get('/api/view/bg')
    def view_bg():
        path = state.get_lyric_bg_image_path()
        if not path or not os.path.isfile(path):
            return JSONResponse(
                {'ok': False, 'message': 'no image'},
                status_code=404,
                headers=_no_store_headers(),
            )
        return FileResponse(path, headers=_no_store_headers())

    @app.get('/view')
    def viewer_page_alias():
        page = static / 'viewer.html'
        if page.is_file():
            return _html_page(page)
        return JSONResponse({'message': 'viewer.html not found'})

    @app.get('/view/text')
    def viewer_text_embed():
        page = static / 'viewer_embed.html'
        if page.is_file():
            return _html_page(page)
        page = static / 'viewer.html'
        if page.is_file():
            return _html_page(page)
        return JSONResponse({'message': 'viewer embed not found'})

    @app.get('/')
    def viewer_page():
        page = static / 'viewer.html'
        if page.is_file():
            return _html_page(page)
        return JSONResponse({'message': 'viewer.html not found'})

    @app.get('/admin')
    def admin_page():
        index = static / 'mobile.html'
        if index.is_file():
            return _html_page(index)
        return JSONResponse({'message': 'mobile.html not found'})

    if static.is_dir():
        app.mount('/static', StaticFiles(directory=str(static)), name='static')

    prefix = (path_prefix or '').strip().rstrip('/')
    if prefix:
        if not prefix.startswith('/'):
            prefix = '/' + prefix
        gateway = FastAPI(title='Hymn Search Remote Gateway')

        @gateway.get(prefix, include_in_schema=False)
        def redirect_prefix_to_slash():
            return RedirectResponse(url=f'{prefix}/', status_code=308)

        # Tunnel: /hymn_search/... ; local LAN: /admin, /api/... (no prefix)
        gateway.mount(prefix, app)
        gateway.mount('/', app)
        return gateway
    return app
