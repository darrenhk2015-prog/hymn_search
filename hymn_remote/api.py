"""FastAPI application for mobile remote control."""
from pathlib import Path

from fastapi import FastAPI, Header, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from .resolver import list_books, list_entries_for_book
from .state import RemoteControlState


class OpenRequest(BaseModel):
    book: str
    num: str = ''


class SetlistOpenRequest(BaseModel):
    index: int


def _static_dir():
    return Path(__file__).resolve().parent / 'static'


def _client_ip(request: Request) -> str:
    if request.client:
        return request.client.host or ''
    return ''


def create_app(state: RemoteControlState) -> FastAPI:
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
    def entries(book: str, q: str = '', x_api_token: str = Header(default='')):
        _check_token(x_api_token)
        return list_entries_for_book(state.get_books(), book, q)

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
        return state.dispatch_open_request(body.book, body.num, _client_ip(request))

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
    def view_now():
        if not state.api_enabled:
            return JSONResponse(
                status_code=503,
                content={'status': 'disabled', 'label': '', 'web_url': '', 'show_debug': False},
            )
        data = state.get_now_viewing()
        data['show_debug'] = state.get_viewer_show_debug()
        return data

    @app.get('/view')
    def viewer_page_alias():
        page = static / 'viewer.html'
        if page.is_file():
            return FileResponse(page)
        return JSONResponse({'message': 'viewer.html not found'})

    @app.get('/')
    def viewer_page():
        page = static / 'viewer.html'
        if page.is_file():
            return FileResponse(page)
        return JSONResponse({'message': 'viewer.html not found'})

    @app.get('/admin')
    def admin_page():
        index = static / 'mobile.html'
        if index.is_file():
            return FileResponse(index)
        return JSONResponse({'message': 'mobile.html not found'})

    if static.is_dir():
        app.mount('/static', StaticFiles(directory=str(static)), name='static')

    return app
