"""测试 Web 版文件匹配功能"""

import urllib.request
import json
import os

BASE = "http://localhost:5890"

def test_scan():
    data = json.dumps({"path": os.path.dirname(os.path.abspath(__file__))}).encode()
    req = urllib.request.Request(f"{BASE}/api/scan_folder", data=data,
        headers={"Content-Type": "application/json"})
    resp = urllib.request.urlopen(req)
    result = json.loads(resp.read())
    print(f"✅ scan_folder: {result['count']} 个文件")
    assert result['count'] > 0
    return result

def test_upload():
    import http.client
    import mimetypes

    fpath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "发送映射模板.xlsx")
    boundary = "----FormBoundary7MA4YWxkTrZu0gW"
    body = b""
    body += f"--{boundary}\r\n".encode()
    body += f'Content-Disposition: form-data; name="file"; filename="{os.path.basename(fpath)}"\r\n'.encode()
    body += b"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n"
    with open(fpath, "rb") as f:
        body += f.read()
    body += f"\r\n--{boundary}--\r\n".encode()

    req = urllib.request.Request(f"{BASE}/api/upload_excel", data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"})
    resp = urllib.request.urlopen(req)
    result = json.loads(resp.read())
    print(f"✅ upload_excel: {result['count']} 条映射 ({result['filename']})")
    assert result['count'] > 0
    return result

def test_preview():
    data = json.dumps({"path": os.path.dirname(os.path.abspath(__file__))}).encode()
    req = urllib.request.Request(f"{BASE}/api/preview", data=data,
        headers={"Content-Type": "application/json"})
    resp = urllib.request.urlopen(req)
    result = json.loads(resp.read())
    print(f"✅ preview: {result['matched_total']} 条映射, {result['sendable']} 可发送, {result['total_files']} 个文件")
    for m in result['matched']:
        print(f"  {m['keyword']} → {m['name']}: {m['files']}")
    return result

def test_status():
    resp = urllib.request.urlopen(f"{BASE}/api/status")
    result = json.loads(resp.read())
    print(f"✅ status: sending={result['sending']}, logs={len(result['logs'])} 条")
    return result

def test_frontend():
    resp = urllib.request.urlopen(f"{BASE}/")
    html = resp.read().decode()
    checks = [
        ("<title>微信文件群发工具</title>", "标题"),
        ("flask", "Flask 框架相关" if "flask" not in html.lower() else ""),
    ]
    for check, desc in checks:
        assert check in html, f"缺少: {desc}"
    print(f"✅ 前端页面加载成功 ({len(html)} bytes)")

if __name__ == "__main__":
    test_scan()
    test_upload()
    test_preview()
    test_status()
    test_frontend()
    print("\n✅ 全部 API 测试通过！")
