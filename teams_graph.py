"""Microsoft Graph SharePoint access for Hangcha Sales Assistant.

Only user-delegated permissions. No password, client secret or refresh token is
saved to disk. All uploads use createUploadSession(conflictBehavior='fail') so
an existing remote PDF is never silently overwritten.
"""
from __future__ import annotations

import base64
from datetime import date
from pathlib import Path
from urllib.parse import quote, urlparse

GRAPH = 'https://graph.microsoft.com/v1.0'
CHUNK = 5 * 1024 * 1024  # 5 MiB = 16 * 320 KiB


class TeamsError(RuntimeError):
    pass


def share_token(url: str) -> str:
    """Official Graph encoding for a share link: u! + base64url UTF-8."""
    url = url.strip()
    parsed = urlparse(url)
    if parsed.scheme != 'https' or not parsed.hostname or not parsed.hostname.lower().endswith('.sharepoint.com'):
        raise TeamsError('请输入公司 SharePoint 的 HTTPS 文件夹链接（:f:）；不要使用 PDF 文件链接。')
    if '/:f:/' not in parsed.path:
        raise TeamsError('分享链接不是文件夹链接：请使用 SharePoint 中“复制链接”获取 :f: 链接。')
    return 'u!' + base64.urlsafe_b64encode(url.encode('utf-8')).decode('ascii').rstrip('=')


def folder_names(day: date) -> tuple[str, str]:
    return f'{day.month}月', day.strftime('%d.%m')


def invoice_candidates(rows, store, profile: str) -> list:
    """Return scanned, eligible invoice rows, including already-uploaded ones.

    Do not filter out uploaded files here: the Teams GUI needs them for the
    '已上传' tab and statistics.  The GUI's _state() separately blocks already
    uploaded, void/credit, duplicate, and unreviewed files from being sent.
    Files no longer present in a scan remain accessible via the history view.
    """
    seen = set()
    output = []
    for row in rows:
        if row.kind != '发票' or row.status in ('忽略作废', '重复文件'):
            continue
        if '作废' in Path(row.source).name:
            continue
        if not row.invoice or not row.contract or not row.source or not Path(row.source).is_file():
            continue
        if row.status not in ('识别成功', '待复核'):
            continue
        # same invoice ID represented by two different files requires review, not auto dedup
        identity = (row.contract, row.invoice, row.sha or str(Path(row.source).resolve()))
        if identity in seen:
            continue
        seen.add(identity)
        output.append(row)
    return output


class GraphClient:
    """Network wrapper is injectable for offline unit tests."""
    def __init__(self, token: str, session=None):
        if not token:
            raise TeamsError('尚未使用公司 Microsoft 365 账户登录')
        if session is None:
            import requests
            session = requests.Session()
        self.token = token
        self.session = session

    def call(self, method, url, *, params=None, json=None, data=None, headers=None, timeout=45, authenticated=True):
        h = {'Accept': 'application/json'}
        if authenticated and url.startswith('https://') and not url.startswith(GRAPH + '/'):
            raise TeamsError('拒绝向非 Microsoft Graph 地址发送公司访问令牌。')
        if authenticated:
            h['Authorization'] = f'Bearer {self.token}'
        if headers:
            h.update(headers)
        response = self.session.request(method, url if url.startswith('https://') else GRAPH + url,
                                        headers=h, params=params, json=json, data=data, timeout=timeout)
        if response.status_code >= 400:
            try:
                err = response.json().get('error', {})
                code, message = err.get('code', ''), err.get('message', '')
            except Exception:
                code, message = '', response.text[:280]
            raise TeamsError(f'Microsoft Graph {response.status_code} {code}: {message}')
        if response.status_code == 204:
            return {}
        return response.json()

    def item(self, drive: str, item_id: str):
        return self.call('GET', f'/drives/{quote(drive,safe="")}/items/{quote(item_id,safe="")}')

    def parent(self, item: dict) -> dict:
        parent = item.get('parentReference') or {}
        if not parent.get('driveId') or not parent.get('id'):
            raise TeamsError('SharePoint 没有返回父文件夹 ID。请改用目标月份的 SharePoint 文件夹链接。')
        return self.item(parent['driveId'], parent['id'])

    def folder_by_share(self, share_url: str):
        found = self.call('GET', '/shares/' + share_token(share_url) + '/driveItem')
        if 'folder' not in found:
            raise TeamsError('共享链接对应的不是文件夹。')
        return found

    def resolve_base(self, share_url: str) -> dict:
        """A date folder link -> parent month -> common months root.
        Month folder link -> its parent. Common root link -> itself.
        The returned base is persisted together with its drive ID and link year.
        """
        obj = self.folder_by_share(share_url)
        name = str(obj.get('name', '')).strip()
        if __import__('re').fullmatch(r'\d{1,2}\.\d{1,2}', name):
            month = self.parent(obj)
            if not __import__('re').fullmatch(r'(?:[1-9]|1[0-2])月', str(month.get('name', ''))):
                raise TeamsError('日期文件夹的上级不是“10月”这样的月份目录，无法自动确定共同根目录。')
            base = self.parent(month)
        elif __import__('re').fullmatch(r'(?:[1-9]|1[0-2])月', name):
            base = self.parent(obj)
        else:
            base = obj
        if 'folder' not in base:
            raise TeamsError('定位得到的目标根目录不是文件夹。')
        parent = base.get('parentReference') or {}
        drive = parent.get('driveId', '')
        if not drive:
            raise TeamsError('无法获取文档库 ID')
        return {'drive_id': drive, 'root_id': base['id'],
                'root_name': base['name'], 'root_url': base.get('webUrl', '')}

    def children(self, drive: str, folder_id: str):
        next_url = f'/drives/{quote(drive,safe="")}/items/{quote(folder_id,safe="")}/children'
        seen_urls = set()
        out = []
        while next_url:
            if next_url in seen_urls:
                raise TeamsError('SharePoint 分页链接重复。')
            seen_urls.add(next_url)
            obj = self.call('GET', next_url)
            out.extend(obj.get('value', []))
            next_url = obj.get('@odata.nextLink')
        return out

    def child(self, drive: str, parent_id: str, exact_name: str) -> dict | None:
        matches = [x for x in self.children(drive, parent_id) if x.get('name') == exact_name and x.get('folder') is not None]
        if len(matches) > 1:
            raise TeamsError('出现重名 SharePoint 目录，停止自动定位。')
        return matches[0] if matches else None

    def target_today(self, base: dict, today: date) -> dict:
        if int(base.get('configured_year', 0)) != today.year:
            raise TeamsError(f'配置为 {base.get("configured_year")} 年，为防止跨年错传，请重新验证目标月份根目录。')
        month_name, day_name = folder_names(today)
        drive = base['drive_id']
        month = self.child(drive, base['root_id'], month_name)
        if not month:
            raise TeamsError(f'找不到月份目录 {month_name}，已停止。请联系文件夹负责人或手动指定目标。')
        day = self.child(drive, month['id'], day_name)
        if not day:
            raise TeamsError(f'找不到日期目录 {month_name}/{day_name}，已停止。不会上传到其他日期。')
        return {'drive_id': drive, 'folder_id': day['id'], 'folder_url': day.get('webUrl', ''),
                'label': f'{base["root_name"]}/{month_name}/{day_name}', 'day': today.isoformat(),
                'manual': False}

    def explicit_target(self, share_url: str, today: date) -> dict:
        folder = self.folder_by_share(share_url)
        drive = (folder.get('parentReference') or {}).get('driveId')
        if not drive:
            raise TeamsError('手动文件夹缺少 Drive ID。')
        return {'drive_id': drive, 'folder_id': folder['id'], 'folder_url': folder.get('webUrl',''),
                'label': '手动选择 / ' + folder['name'], 'day': today.isoformat(), 'manual': True}

    def remote_file(self, drive: str, folder_id: str, name: str) -> dict | None:
        # Query children rather than build an opaque path. Unicode and '#' are safe.
        matches = [x for x in self.children(drive, folder_id) if x.get('name', '').casefold() == name.casefold()]
        if len(matches) > 1:
            raise TeamsError(f'远程存在多个大小写相同的文件名：{name}')
        return matches[0] if matches else None

    def upload_no_overwrite(self, file: Path, target: dict) -> dict:
        """No overwrite even if remote file appears after preview.
        Uses createUploadSession(conflictBehavior=fail), with 5 MiB aligned chunks.
        """
        if not file.is_file() or file.suffix.lower() != '.pdf' or file.stat().st_size == 0:
            raise TeamsError('本地发票不存在、为空或不是 PDF')
        drive, folder_id = target['drive_id'], target['folder_id']
        filename = quote(file.name, safe='')
        session = self.call('POST', f'/drives/{quote(drive,safe="")}/items/{quote(folder_id,safe="")}:/'+filename+':/createUploadSession',
                            json={'item': {'@microsoft.graph.conflictBehavior': 'fail', 'name': file.name},
                                  'deferCommit': False})
        upload_url = session.get('uploadUrl')
        if not upload_url or not upload_url.startswith('https://'):
            raise TeamsError('服务器没有返回有效的 Upload Session URL')
        size = file.stat().st_size
        result = None
        with file.open('rb') as f:
            sent = 0
            while sent < size:
                chunk = f.read(CHUNK)
                if not chunk:
                    raise TeamsError('读取 PDF 时意外中断')
                result = self.call('PUT', upload_url, authenticated=False, data=chunk,
                                   headers={'Content-Length': str(len(chunk)),
                                            'Content-Range': f'bytes {sent}-{sent+len(chunk)-1}/{size}',
                                            'Content-Type': 'application/octet-stream'}, timeout=180)
                sent += len(chunk)
        if not result or not result.get('id'):
            raise TeamsError('上传已发送，但没有收到 SharePoint 文件 ID：上传状态未知，请在 Teams 核查后重试。')
        if result.get('size') != size:
            raise TeamsError('上传后的文件大小与本地 PDF 不一致，请人工检查。')
        return result
