//go:build windows

package excel

import (
	"bytes"
	"fmt"
	"net"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"sync/atomic"
	"time"
	"unsafe"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"golang.org/x/sys/windows"
	"golang.org/x/sys/windows/registry"
)

// xlOpenXMLWorkbook is Excel's FileFormat for .xlsx (no macros).
const xlOpenXMLWorkbook int32 = 51

func newPlatformHost() Host {
	return &windowsHost{}
}

type windowsHost struct {
	mu   sync.Mutex
	last HostInfo
}

func (h *windowsHost) Available() bool {
	_, err := ole.ClassIDFrom("Excel.Application")
	return err == nil
}

func (h *windowsHost) Info() (HostInfo, error) {
	h.mu.Lock()
	last := h.last
	h.mu.Unlock()
	if last.ExcelVersion != "" {
		return last, nil
	}
	info, err := queryExcelInfo(DefaultTimeout)
	if err != nil {
		return HostInfo{ID: HostID, OS: runtime.GOOS}, err
	}
	h.mu.Lock()
	h.last = info
	h.mu.Unlock()
	return info, nil
}

func (h *windowsHost) OpenSave(inputPath, outputPath string) error {
	absIn, absOut, err := preparePaths(inputPath, outputPath)
	if err != nil {
		return err
	}
	info, err := runExcelSTA(DefaultTimeout, func(excel *ole.IDispatch) (HostInfo, error) {
		return openSaveWithExcel(excel, absIn, absOut)
	})
	if err != nil {
		return err
	}
	h.mu.Lock()
	h.last = info
	h.mu.Unlock()
	return nil
}

func (h *windowsHost) RunScript(inputPath, scriptPath, outputPath string) error {
	absIn, absScript, absOut, err := prepareScriptRun(inputPath, scriptPath, outputPath)
	if err != nil {
		return err
	}
	scriptBytes, err := os.ReadFile(absScript)
	if err != nil {
		return err
	}
	if len(bytes.TrimSpace(scriptBytes)) == 0 {
		return fmt.Errorf("script is empty; use excel-save for load+save")
	}

	job := newJobServer(string(scriptBytes))
	baseURL, closeSrv, err := serveAddin(job)
	if err != nil {
		return err
	}
	defer closeSrv()

	catalogDir, err := PersistentCatalogDir()
	if err != nil {
		return err
	}
	ver := sideloadVersion()
	if err := writeSideloadCatalog(catalogDir, baseURL, ver); err != nil {
		return err
	}
	if err := registerSideload(catalogDir); err != nil {
		return err
	}
	runtimeLog := filepath.Join(catalogDir, "runtime.log")
	_ = os.Remove(runtimeLog)
	if err := enableRuntimeLogging(runtimeLog); err != nil {
		return err
	}
	stamped := filepath.Join(catalogDir, "stamped.xlsx")
	if err := stampWebExtension(stamped, absIn, AddinID, ver); err != nil {
		return err
	}

	scriptErr := make(chan error, 1)
	go func() {
		_, err := job.wait(DefaultTimeout)
		scriptErr <- err
	}()

	// COM CreateObject Excel does not load Office.js / WebView2. Start excel.exe
	// as a real process (same as a user double-click) and attach to SaveAs.
	pid, err := startExcelProcess(stamped)
	if err != nil {
		return err
	}
	info, err := runAttachedExcel(pid, DefaultTimeout+15*time.Second, func(excel *ole.IDispatch) (HostInfo, error) {
		return saveAfterScript(excel, absOut, scriptErr)
	})
	if err != nil {
		return fmt.Errorf("%w (%s)", err, runtimeLogTail(runtimeLog))
	}
	h.mu.Lock()
	h.last = info
	h.mu.Unlock()
	return nil
}

func queryExcelInfo(timeout time.Duration) (HostInfo, error) {
	return runExcelSTA(timeout, func(excel *ole.IDispatch) (HostInfo, error) {
		return readExcelInfo(excel), nil
	})
}

type excelJob func(excel *ole.IDispatch) (HostInfo, error)

func runExcelSTA(timeout time.Duration, job excelJob) (HostInfo, error) {
	type result struct {
		info HostInfo
		err  error
	}
	done := make(chan result, 1)
	var pid atomic.Uint32

	go func() {
		runtime.LockOSThread()
		defer runtime.UnlockOSThread()
		info, err := withExcelApplication(&pid, job)
		done <- result{info: info, err: err}
	}()

	timer := time.NewTimer(timeout)
	defer timer.Stop()
	select {
	case r := <-done:
		return r.info, r.err
	case <-timer.C:
		if p := pid.Load(); p != 0 {
			killPID(p)
		}
		return HostInfo{}, fmt.Errorf("excel-save timed out after %s", timeout)
	}
}

func withExcelApplication(pid *atomic.Uint32, job excelJob) (info HostInfo, err error) {
	if e := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED); e != nil {
		// S_FALSE (already initialized) is reported as error by some wrappers;
		// continue if CreateObject still works.
	}
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("Excel.Application")
	if err != nil {
		return HostInfo{ID: HostID, OS: runtime.GOOS},
			fmt.Errorf("Excel.Application: %w (is Microsoft Excel installed?)", err)
	}
	defer unknown.Release()

	excel, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return HostInfo{ID: HostID, OS: runtime.GOOS}, fmt.Errorf("Excel IDispatch: %w", err)
	}
	defer excel.Release()

	if p := pidFromExcel(excel); p != 0 {
		pid.Store(p)
	}

	defer func() {
		_, _ = oleutil.CallMethod(excel, "Quit")
		if p := pid.Load(); p != 0 {
			waitOrKill(p, 10*time.Second)
		}
	}()

	if _, err := oleutil.PutProperty(excel, "Visible", false); err != nil {
		return info, fmt.Errorf("Excel Visible=false: %w", err)
	}
	if _, err := oleutil.PutProperty(excel, "DisplayAlerts", false); err != nil {
		return info, fmt.Errorf("Excel DisplayAlerts=false: %w", err)
	}
	_, _ = oleutil.PutProperty(excel, "AskToUpdateLinks", false)
	_, _ = oleutil.PutProperty(excel, "AlertBeforeOverwriting", false)
	_, _ = oleutil.PutProperty(excel, "ScreenUpdating", false)

	return job(excel)
}

func openSaveWithExcel(excel *ole.IDispatch, absIn, absOut string) (HostInfo, error) {
	info := readExcelInfo(excel)

	wbProp, err := oleutil.GetProperty(excel, "Workbooks")
	if err != nil {
		return info, fmt.Errorf("Excel Workbooks: %w", err)
	}
	defer wbProp.Clear()
	workbooks := wbProp.ToIDispatch()
	if workbooks == nil {
		return info, fmt.Errorf("Excel Workbooks is nil")
	}

	// UpdateLinks=0 (don't update), ReadOnly=false. int32 so COM gets VT_I4.
	opened, err := oleutil.CallMethod(workbooks, "Open", absIn, int32(0), false)
	if err != nil {
		return info, fmt.Errorf("Excel Workbooks.Open %s: %w", absIn, err)
	}
	defer opened.Clear()
	wb := opened.ToIDispatch()
	if wb == nil {
		return info, fmt.Errorf("Excel Workbooks.Open returned nil for %s", absIn)
	}
	defer func() {
		_, _ = oleutil.CallMethod(wb, "Close", false)
	}()

	if sameFilePath(absIn, absOut) {
		if _, err := oleutil.CallMethod(wb, "Save"); err != nil {
			return info, fmt.Errorf("Excel Workbook.Save: %w", err)
		}
	} else {
		if _, err := oleutil.CallMethod(wb, "SaveAs", absOut, xlOpenXMLWorkbook); err != nil {
			return info, fmt.Errorf("Excel Workbook.SaveAs %s: %w", absOut, err)
		}
	}
	return info, nil
}

func saveAfterScript(excel *ole.IDispatch, absOut string, scriptErr <-chan error) (HostInfo, error) {
	info := readExcelInfo(excel)
	_, _ = oleutil.PutProperty(excel, "ScreenUpdating", true)
	_, _ = oleutil.PutProperty(excel, "EnableEvents", true)
	_, _ = oleutil.PutProperty(excel, "AutomationSecurity", int32(1))

	// Wait longer than job.wait so its detailed timeout (HTTP hits) wins the race.
	if err := pumpUntil(scriptErr, DefaultTimeout+5*time.Second); err != nil {
		return info, err
	}

	if _, err := oleutil.PutProperty(excel, "DisplayAlerts", false); err != nil {
		return info, fmt.Errorf("Excel DisplayAlerts=false: %w", err)
	}
	wb, err := activeWorkbook(excel)
	if err != nil {
		return info, err
	}
	defer func() {
		_, _ = oleutil.CallMethod(wb, "Close", false)
	}()
	if _, err := oleutil.CallMethod(wb, "SaveAs", absOut, xlOpenXMLWorkbook); err != nil {
		return info, fmt.Errorf("Excel Workbook.SaveAs %s: %w", absOut, err)
	}
	return info, nil
}

func activeWorkbook(excel *ole.IDispatch) (*ole.IDispatch, error) {
	v, err := oleutil.GetProperty(excel, "ActiveWorkbook")
	if err != nil {
		return nil, fmt.Errorf("Excel ActiveWorkbook: %w", err)
	}
	defer v.Clear()
	wb := v.ToIDispatch()
	if wb == nil {
		return nil, fmt.Errorf("Excel ActiveWorkbook is nil")
	}
	wb.AddRef()
	return wb, nil
}

func pumpUntil(done <-chan error, timeout time.Duration) error {
	deadline := time.Now().Add(timeout)
	for {
		pumpMessages()
		select {
		case err := <-done:
			return err
		default:
		}
		if time.Now().After(deadline) {
			return fmt.Errorf("excel-run timed out waiting for Office.js after %s", timeout)
		}
		time.Sleep(15 * time.Millisecond)
	}
}

func pumpMessages() {
	var msg [64]byte
	for {
		r, _, _ := peekMessage.Call(uintptr(unsafe.Pointer(&msg[0])), 0, 0, 0, 1)
		if r == 0 {
			return
		}
		_, _, _ = translateMessage.Call(uintptr(unsafe.Pointer(&msg[0])))
		_, _, _ = dispatchMessage.Call(uintptr(unsafe.Pointer(&msg[0])))
	}
}

func readExcelInfo(excel *ole.IDispatch) HostInfo {
	info := HostInfo{ID: HostID, OS: runtime.GOOS}
	if v, err := oleutil.GetProperty(excel, "Version"); err == nil {
		info.ExcelVersion = variantString(v)
		_ = v.Clear()
	}
	if v, err := oleutil.GetProperty(excel, "Build"); err == nil {
		info.ExcelBuild = variantString(v)
		_ = v.Clear()
	}
	return info
}

func variantString(v *ole.VARIANT) string {
	if v == nil {
		return ""
	}
	if s := v.ToString(); s != "" {
		return s
	}
	if val := v.Value(); val != nil {
		return fmt.Sprint(val)
	}
	return ""
}

func pidFromExcel(excel *ole.IDispatch) uint32 {
	v, err := oleutil.GetProperty(excel, "Hwnd")
	if err != nil {
		return 0
	}
	defer v.Clear()
	hwnd := uintptr(v.Val)
	if hwnd == 0 {
		return 0
	}
	var pid uint32
	getWindowThreadProcessId.Call(hwnd, uintptr(unsafe.Pointer(&pid)))
	return pid
}

var (
	user32                     = windows.NewLazySystemDLL("user32.dll")
	oleacc                     = windows.NewLazySystemDLL("oleacc.dll")
	getWindowThreadProcessId   = user32.NewProc("GetWindowThreadProcessId")
	peekMessage                = user32.NewProc("PeekMessageW")
	translateMessage           = user32.NewProc("TranslateMessage")
	dispatchMessage            = user32.NewProc("DispatchMessageW")
	findWindowEx               = user32.NewProc("FindWindowExW")
	accessibleObjectFromWindow = oleacc.NewProc("AccessibleObjectFromWindow")
)

const objidNativeOM = 0xFFFFFFF0

func serveAddin(handler http.Handler) (baseURL string, closeFn func(), err error) {
	ln4, err := net.Listen("tcp4", "127.0.0.1:0")
	if err != nil {
		return "", nil, fmt.Errorf("add-in server: %w", err)
	}
	port := ln4.Addr().(*net.TCPAddr).Port
	lns := []net.Listener{ln4}
	if ln6, err6 := net.Listen("tcp6", fmt.Sprintf("[::1]:%d", port)); err6 == nil {
		lns = append(lns, ln6)
	}
	for _, ln := range lns {
		go func(ln net.Listener) { _ = http.Serve(ln, handler) }(ln)
	}
	return fmt.Sprintf("http://localhost:%d", port), func() {
		for _, ln := range lns {
			_ = ln.Close()
		}
	}, nil
}

func startExcelProcess(xlsx string) (uint32, error) {
	exe, err := excelExecutable()
	if err != nil {
		return 0, err
	}
	cmd := exec.Command(exe, "/x", xlsx)
	if err := cmd.Start(); err != nil {
		return 0, fmt.Errorf("start Excel: %w", err)
	}
	return uint32(cmd.Process.Pid), nil
}

func excelExecutable() (string, error) {
	hives := []registry.Key{registry.CURRENT_USER, registry.LOCAL_MACHINE}
	subs := []string{
		`SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe`,
		`SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe`,
	}
	for _, hive := range hives {
		for _, sub := range subs {
			k, err := registry.OpenKey(hive, sub, registry.QUERY_VALUE)
			if err != nil {
				continue
			}
			path, _, err := k.GetStringValue("")
			k.Close()
			if err != nil || path == "" {
				continue
			}
			if _, err := os.Stat(path); err == nil {
				return path, nil
			}
		}
	}
	for _, c := range []string{
		`C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE`,
		`C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE`,
		`C:\Program Files\Microsoft Office\Office16\EXCEL.EXE`,
	} {
		if _, err := os.Stat(c); err == nil {
			return c, nil
		}
	}
	return "", fmt.Errorf("excel.exe not found")
}

func runAttachedExcel(startPID uint32, timeout time.Duration, job excelJob) (HostInfo, error) {
	type result struct {
		info HostInfo
		err  error
	}
	done := make(chan result, 1)
	var livePID atomic.Uint32
	livePID.Store(startPID)

	go func() {
		runtime.LockOSThread()
		defer runtime.UnlockOSThread()
		if e := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED); e != nil {
		}
		defer ole.CoUninitialize()

		excel, err := waitAttachExcel(startPID, 45*time.Second)
		if err != nil {
			done <- result{info: HostInfo{ID: HostID, OS: runtime.GOOS}, err: err}
			return
		}
		if p := pidFromExcel(excel); p != 0 {
			livePID.Store(p)
		}
		defer func() {
			_, _ = oleutil.CallMethod(excel, "Quit")
			excel.Release()
			if p := livePID.Load(); p != 0 {
				waitOrKill(p, 10*time.Second)
			}
		}()
		info, err := job(excel)
		done <- result{info: info, err: err}
	}()

	timer := time.NewTimer(timeout)
	defer timer.Stop()
	select {
	case r := <-done:
		return r.info, r.err
	case <-timer.C:
		if p := livePID.Load(); p != 0 {
			killPID(p)
		}
		return HostInfo{}, fmt.Errorf("excel-run timed out after %s", timeout)
	}
}

func waitAttachExcel(startPID uint32, timeout time.Duration) (*ole.IDispatch, error) {
	deadline := time.Now().Add(timeout)
	var last error
	for time.Now().Before(deadline) {
		pumpMessages()
		disp, err := attachExcel(startPID)
		if err == nil {
			return disp, nil
		}
		last = err
		time.Sleep(100 * time.Millisecond)
	}
	if last == nil {
		last = fmt.Errorf("no Excel window")
	}
	return nil, fmt.Errorf("attach Excel: %w", last)
}

func attachExcel(startPID uint32) (*ole.IDispatch, error) {
	pids := descendantPIDs(startPID)
	xlmain := findXLMAIN(pids)
	if xlmain == 0 {
		return nil, fmt.Errorf("no XLMAIN for pid %v", pids)
	}
	desk := findChild(xlmain, 0, "XLDESK")
	if desk == 0 {
		return nil, fmt.Errorf("no XLDESK")
	}
	book := findChild(desk, 0, "EXCEL7")
	if book == 0 {
		return nil, fmt.Errorf("no EXCEL7")
	}
	var win *ole.IDispatch
	hr, _, _ := accessibleObjectFromWindow.Call(
		book,
		objidNativeOM,
		uintptr(unsafe.Pointer(ole.IID_IDispatch)),
		uintptr(unsafe.Pointer(&win)),
	)
	if hr != 0 || win == nil {
		if app, err := activeExcelMatching(pids); err == nil {
			return app, nil
		}
		return nil, fmt.Errorf("AccessibleObjectFromWindow hr=0x%x", hr)
	}
	appVar, err := oleutil.GetProperty(win, "Application")
	win.Release()
	if err != nil {
		return nil, fmt.Errorf("Window.Application: %w", err)
	}
	defer appVar.Clear()
	app := appVar.ToIDispatch()
	if app == nil {
		return nil, fmt.Errorf("Window.Application is nil")
	}
	app.AddRef()
	return app, nil
}

func activeExcelMatching(pids []uint32) (*ole.IDispatch, error) {
	unk, err := oleutil.GetActiveObject("Excel.Application")
	if err != nil {
		return nil, err
	}
	defer unk.Release()
	app, err := unk.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return nil, err
	}
	got := pidFromExcel(app)
	for _, p := range pids {
		if p != 0 && p == got {
			return app, nil
		}
	}
	app.Release()
	return nil, fmt.Errorf("GetActiveObject pid %d not in %v", got, pids)
}

func findChild(parent, after uintptr, class string) uintptr {
	cls, err := windows.UTF16PtrFromString(class)
	if err != nil {
		return 0
	}
	hwnd, _, _ := findWindowEx.Call(parent, after, uintptr(unsafe.Pointer(cls)), 0)
	return hwnd
}

func findXLMAIN(pids []uint32) uintptr {
	want := map[uint32]bool{}
	for _, p := range pids {
		want[p] = true
	}
	cls, err := windows.UTF16PtrFromString("XLMAIN")
	if err != nil {
		return 0
	}
	var hwnd uintptr
	for {
		next, _, _ := findWindowEx.Call(0, hwnd, uintptr(unsafe.Pointer(cls)), 0)
		if next == 0 {
			return 0
		}
		hwnd = next
		var wpid uint32
		getWindowThreadProcessId.Call(hwnd, uintptr(unsafe.Pointer(&wpid)))
		if want[wpid] {
			return hwnd
		}
	}
}

func descendantPIDs(root uint32) []uint32 {
	children := childrenByParent()
	out := []uint32{root}
	queue := []uint32{root}
	seen := map[uint32]bool{root: true}
	for len(queue) > 0 {
		p := queue[0]
		queue = queue[1:]
		for _, c := range children[p] {
			if seen[c] {
				continue
			}
			seen[c] = true
			out = append(out, c)
			queue = append(queue, c)
		}
	}
	return out
}

func childrenByParent() map[uint32][]uint32 {
	snap, err := windows.CreateToolhelp32Snapshot(windows.TH32CS_SNAPPROCESS, 0)
	if err != nil {
		return nil
	}
	defer windows.CloseHandle(snap)
	var e windows.ProcessEntry32
	e.Size = uint32(unsafe.Sizeof(e))
	if err := windows.Process32First(snap, &e); err != nil {
		return nil
	}
	out := map[uint32][]uint32{}
	for {
		out[e.ParentProcessID] = append(out[e.ParentProcessID], e.ProcessID)
		if err := windows.Process32Next(snap, &e); err != nil {
			break
		}
	}
	return out
}

func waitOrKill(pid uint32, wait time.Duration) {
	deadline := time.Now().Add(wait)
	for time.Now().Before(deadline) {
		if !pidAlive(pid) {
			return
		}
		time.Sleep(100 * time.Millisecond)
	}
	killPID(pid)
}

func pidAlive(pid uint32) bool {
	const stillActive = 259 // STILL_ACTIVE
	h, err := windows.OpenProcess(windows.PROCESS_QUERY_LIMITED_INFORMATION, false, pid)
	if err != nil {
		return false
	}
	defer windows.CloseHandle(h)
	var code uint32
	if err := windows.GetExitCodeProcess(h, &code); err != nil {
		return false
	}
	return code == stillActive
}

func killPID(pid uint32) {
	proc, err := os.FindProcess(int(pid))
	if err != nil {
		return
	}
	_ = proc.Kill()
}

func runtimeLogTail(path string) string {
	data, err := os.ReadFile(path)
	if err != nil {
		return "no Office add-in runtime log"
	}
	s := strings.TrimSpace(string(data))
	if s == "" {
		return "Office add-in runtime log empty"
	}
	const max = 2000
	if len(s) > max {
		s = s[len(s)-max:]
	}
	return "runtime.log: " + s
}
