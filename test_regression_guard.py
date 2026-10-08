import os, sys, re, json, time, subprocess

sys.stdout.reconfigure(encoding='utf-8')

CHROME_CANDIDATE_PATHS = [
    r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
    r'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
    r'C:\Program Files\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    os.path.expandvars(r'%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe'),
    os.path.expandvars(r'%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe'),
]

browser_exe = None
for p in CHROME_CANDIDATE_PATHS:
    if os.path.exists(p):
        browser_exe = p
        break

if not browser_exe:
    print("ERROR: Không tìm thấy trình duyệt msedge hoặc chrome trên hệ thống.")
    sys.exit(1)

test_url = sys.argv[1] if (len(sys.argv) > 1 and sys.argv[1].startswith("http")) else "http://127.0.0.1:8080/index.html?run_guard=1"
print(f"================================================================================")
print(f"   HỆ THỐNG KIỂM THỬ TỰ ĐỘNG BẢO VỆ CHỐNG TÁI PHÁT LỖI (REGRESSION GUARD)     ")
print(f"================================================================================")
print(f"Trình duyệt: {browser_exe}")
print(f"Địa chỉ test: {test_url}\n")

# Run headless browser with dump-dom and virtual time budget to allow execution
cmd = [
    browser_exe,
    '--headless=new',
    '--disable-gpu',
    '--no-sandbox',
    '--run-all-compositor-stages-before-draw',
    '--virtual-time-budget=8000',
    '--dump-dom',
    test_url
]

try:
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', timeout=20)
    stdout = proc.stdout
except Exception as e:
    print(f"Lỗi chạy browser headless: {e}")
    sys.exit(1)

m = re.search(r'<pre\s+id=["\']guard-output["\']>([\s\S]*?)</pre>', stdout)
if not m:
    print("CẢNH BÁO: Chưa tìm thấy kết quả trong DOM! Thử phân tích console log...")
    # fallback search for json in DOM
    m2 = re.search(r'("all_passed":\s*(?:true|false)[\s\S]*?\{[\s\S]*?\})', stdout)
    if not m2:
        print("DOM trích xuất không chứa guard-output. Nội dung DOM rút gọn:")
        print(stdout[-1500:])
        sys.exit(1)
    json_text = m2.group(1)
else:
    json_text = m.group(1).strip()

try:
    data = json.loads(json_text)
except Exception as err:
    print(f"Lỗi parse JSON kết quả: {err}\nChuỗi: {json_text[:300]}")
    sys.exit(1)

total = data.get('total', 0)
passed = data.get('passed', 0)
failed = data.get('failed', 0)
rate = data.get('pass_rate', '0%')
all_ok = data.get('all_passed', False)

print(f"{'ID':<4} | {'TÌNH HUỐNG THỬ NGHIỆM':<50} | {'KỲ VỌNG':<12} | {'THỰC TẾ':<12} | {'TRẠNG THÁI'}")
print("-" * 95)

for tc in data.get('cases', []):
    status_str = "✅ PASS" if tc.get('passed') else "❌ FAIL"
    title_short = tc.get('title', '')[:48]
    exp = tc.get('expected', '')
    act = tc.get('actual', '')
    print(f"{tc.get('id'):<4} | {title_short:<50} | {exp:<12} | {act:<12} | {status_str}")

print("-" * 95)
print(f"TỔNG KẾT: {passed}/{total} ca đạt chuẩn ({rate})")

if all_ok:
    print("\n🎉 XÁC NHẬN: 100% CÁC TÌNH HUỐNG BENCHMARK PHÁP LÝ HẢI QUAN ĐỀU VƯỢT QUA!")
    print("Hệ thống an toàn tuyệt đối, không có bất kỳ hiện tượng tái phát lỗi (Zero Regression).\n")
    sys.exit(0)
else:
    print(f"\n⚠️ CẢNH BÁO: CÓ {failed} TÌNH HUỐNG KHÔNG ĐẠT YÊU CẦU! VUI LÒNG KIỂM TRA LẠI.")
    sys.exit(1)
