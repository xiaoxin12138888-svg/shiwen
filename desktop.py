"""Per-user Windows entry point; no external Python or Node installation needed."""
import ctypes
import json
import msvcrt
import os
import sys
import time
import traceback
import urllib.request
import webbrowser
from http.server import ThreadingHTTPServer
import app

INFO=app.DATA/'desktop-session.json'

def existing():
    try:
        info=json.loads(INFO.read_text('utf-8'))
        port=int(info['port'])
        if not 1 <= port <= 65535:return None
        url=f'http://127.0.0.1:{port}'
        with urllib.request.urlopen(url+'/health',timeout=1) as r:health=json.load(r)
        if health.get('desktop') and health.get('instance')==info['token'][:12]:
            return url,info['token']
    except (OSError,ValueError,KeyError):pass
    return None

def main():
    if '--shutdown' in sys.argv:
        current=existing()
        if current:
            url,token=current
            request=urllib.request.Request(url+'/api/shutdown',data=b'{"force":true}',headers={'Content-Type':'application/json','X-Collector-Token':token})
            with urllib.request.urlopen(request,timeout=5) as response:response.read()
            for _ in range(50):
                if not existing():break
                time.sleep(.1)
        return
    lock=open(app.DATA/'desktop.lock','a+b')
    if lock.tell()==0:lock.write(b'0');lock.flush()
    lock.seek(0)
    try:msvcrt.locking(lock.fileno(),msvcrt.LK_NBLCK,1)
    except OSError:
        for _ in range(40):
            current=existing()
            if current:
                if '--no-browser' not in sys.argv:webbrowser.open(current[0])
                return
            time.sleep(.1)
        raise RuntimeError('拾文正在启动或退出，请稍等片刻后重试。')
    try:
        app.initialize()
        preferred=int(os.environ.get('COLLECTOR_PORT','18766'))
        try:server=ThreadingHTTPServer(('127.0.0.1',preferred),app.Handler)
        except OSError:server=ThreadingHTTPServer(('127.0.0.1',0),app.Handler)
        app.PORT=server.server_port
        temp=INFO.with_suffix('.tmp')
        temp.write_text(json.dumps({'port':app.PORT,'token':app.TOKEN}),encoding='utf-8')
        os.replace(temp,INFO)
        if '--no-browser' not in sys.argv:webbrowser.open(f'http://127.0.0.1:{app.PORT}')
        try:server.serve_forever(poll_interval=.2)
        finally:server.server_close();INFO.unlink(missing_ok=True)
    finally:
        lock.seek(0);msvcrt.locking(lock.fileno(),msvcrt.LK_UNLCK,1);lock.close()

if __name__=='__main__':
    try:main()
    except Exception:
        (app.DATA/'startup-error.log').write_text(traceback.format_exc(),encoding='utf-8')
        ctypes.windll.user32.MessageBoxW(0,'拾文未能启动。请重试，或查看数据目录中的 startup-error.log。\n\n'+str(app.DATA),'拾文',0x10)
