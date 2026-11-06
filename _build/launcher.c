/*
 * 客户Excel处理工具 —— 便携包启动器（原生 C 版）
 *
 * 作用与 launcher.py 完全相同，但：
 *   - 编译出来约 140 KB（/MT 静态链接 CRT），只依赖系统 DLL，不需要 Python 运行时
 *   - 不做自解压，没有 --onefile 那种"运行时释放临时文件"的查杀特征
 *   - 没有任何第三方导入，静态启发式很难判成可疑程序
 *
 * 布局约定：
 *   <根目录>/客户Excel处理工具.exe   <- 本启动器（唯一可见的 exe）
 *   <根目录>/_runtime/新需求.exe      <- Nuitka standalone 编译的真实主程序
 *
 * 启动器负责：设置 FAPIAO_ROOT_DIR、把 _runtime 与 _runtime\lib 前置进 PATH、
 * 以根目录为工作目录启动主程序，并透传退出码。
 *
 * 编译方式见 _build/build_portable.py（rc + cl /O2 /MT /utf-8）。
 */

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include <windows.h>
#include <strsafe.h>
#include <stdarg.h>
#include <stdio.h>
#include <wchar.h>

#define RUNTIME_DIR_NAME L"_runtime"
#define MAIN_EXE_NAME    L"新需求.exe"  /* 新需求.exe */
#define ROOT_DIR_ENV     L"FAPIAO_ROOT_DIR"
#define BUF_CHARS        32768

static wchar_t g_exe_path[BUF_CHARS];
static wchar_t g_root[BUF_CHARS];
static wchar_t g_runtime[BUF_CHARS];
static wchar_t g_main_exe[BUF_CHARS];
static wchar_t g_cmdline[BUF_CHARS * 2];

/* ---------------- 控制台输出 ---------------- */
/* 中文必须走 WriteConsoleW：wprintf 在默认 locale 下会把中文变成 '?'。
   输出被重定向到文件/管道时改用 UTF-8，方便自动化抓日志。 */
static void write_std(HANDLE handle, const wchar_t *text)
{
    DWORD written = 0;
    int need;
    char stack_buf[1024];
    char *buf = stack_buf;
    char *heap = NULL;
    int got;

    if (handle == INVALID_HANDLE_VALUE || text == NULL)
        return;

    if (GetFileType(handle) == FILE_TYPE_CHAR) {
        WriteConsoleW(handle, text, (DWORD)wcslen(text), &written, NULL);
        return;
    }

    need = WideCharToMultiByte(CP_UTF8, 0, text, -1, NULL, 0, NULL, NULL);
    if (need <= 1)
        return;
    if (need > (int)sizeof(stack_buf)) {
        heap = (char *)HeapAlloc(GetProcessHeap(), 0, (SIZE_T)need);
        if (heap == NULL)
            return;
        buf = heap;
    }
    got = WideCharToMultiByte(CP_UTF8, 0, text, -1, buf, need, NULL, NULL);
    if (got > 0)
        WriteFile(handle, buf, (DWORD)(got - 1), &written, NULL);
    if (heap != NULL)
        HeapFree(GetProcessHeap(), 0, heap);
}

static void out_text(const wchar_t *text)
{
    write_std(GetStdHandle(STD_OUTPUT_HANDLE), text);
}

static void out_fmt(const wchar_t *fmt, ...)
{
    wchar_t buf[2048];
    va_list ap;

    va_start(ap, fmt);
    _vsnwprintf_s(buf, _countof(buf), _TRUNCATE, fmt, ap);
    va_end(ap);
    out_text(buf);
}

static void press_any_key(void)
{
    HANDLE in = GetStdHandle(STD_INPUT_HANDLE);
    DWORD mode = 0;
    wchar_t ch = 0;
    DWORD read = 0;

    out_text(L"\n按任意键关闭窗口 ……");
    if (in != INVALID_HANDLE_VALUE && GetConsoleMode(in, &mode))
        ReadConsoleW(in, &ch, 1, &read, NULL);
}

/* ---------------- 路径小工具 ---------------- */
static BOOL path_is_file(const wchar_t *path)
{
    DWORD attr = GetFileAttributesW(path);
    return attr != INVALID_FILE_ATTRIBUTES && (attr & FILE_ATTRIBUTE_DIRECTORY) == 0;
}

static BOOL path_is_dir(const wchar_t *path)
{
    DWORD attr = GetFileAttributesW(path);
    return attr != INVALID_FILE_ATTRIBUTES && (attr & FILE_ATTRIBUTE_DIRECTORY) != 0;
}

static HRESULT join_path(wchar_t *dst, size_t cch, const wchar_t *base, const wchar_t *name)
{
    size_t n = wcslen(base);

    if (n > 0 && (base[n - 1] == L'\\' || base[n - 1] == L'/'))
        return StringCchPrintfW(dst, cch, L"%s%s", base, name);
    return StringCchPrintfW(dst, cch, L"%s\\%s", base, name);
}

/* 由启动器自身路径推出"操作员可见的根目录" */
static BOOL resolve_root(void)
{
    DWORD len = GetModuleFileNameW(NULL, g_exe_path, BUF_CHARS);
    wchar_t *slash;
    size_t n;

    if (len == 0 || len >= BUF_CHARS)
        return FALSE;
    if (FAILED(StringCchCopyW(g_root, BUF_CHARS, g_exe_path)))
        return FALSE;

    slash = wcsrchr(g_root, L'\\');
    if (slash == NULL) {
        return SUCCEEDED(StringCchCopyW(g_root, BUF_CHARS, L"."));
    }

    *slash = L'\0';
    n = wcslen(g_root);
    /* "D:" 这种盘符根目录要补反斜杠，否则 "D:_runtime" 会被解析成相对路径 */
    if (n > 0 && g_root[n - 1] == L':')
        return SUCCEEDED(StringCchCatW(g_root, BUF_CHARS, L"\\"));
    return TRUE;
}

/* 取 GetCommandLineW() 里 exe 之后的原始参数（保留引号原样） */
static const wchar_t *args_after_exe(void)
{
    const wchar_t *p = GetCommandLineW();

    if (p == NULL)
        return L"";

    if (*p == L'"') {
        ++p;
        while (*p != L'\0' && *p != L'"')
            ++p;
        if (*p == L'"')
            ++p;
    } else {
        while (*p != L'\0' && *p != L' ' && *p != L'\t')
            ++p;
    }
    while (*p == L' ' || *p == L'\t')
        ++p;
    return p;
}

/* 解压后隐藏属性可能丢失（zip/资源管理器不保证保留 DOS 属性），
   每次启动顺手补一次，保证操作员看到的是"根目录只有一个 exe"。 */
static void ensure_runtime_hidden(void)
{
    DWORD attr = GetFileAttributesW(g_runtime);

    if (attr == INVALID_FILE_ATTRIBUTES)
        return;
    if (attr & FILE_ATTRIBUTE_HIDDEN)
        return;
    SetFileAttributesW(g_runtime, attr | FILE_ATTRIBUTE_HIDDEN);
}

static void prepare_environment(void)
{
    wchar_t lib_dir[BUF_CHARS];
    wchar_t old_path[BUF_CHARS];
    wchar_t new_path[BUF_CHARS * 2];
    DWORD old_len;

    /* 1. 告诉主程序"根目录"在哪（1.py 里 _app_base_dir() 第一优先级读它） */
    SetEnvironmentVariableW(ROOT_DIR_ENV, g_root);
    /* 2. 把 _runtime 与 _runtime\lib 前置进 PATH，辅助 DLL 加载 */
    join_path(lib_dir, _countof(lib_dir), g_runtime, L"lib");

    old_path[0] = L'\0';
    old_len = GetEnvironmentVariableW(L"PATH", old_path, BUF_CHARS);
    if (old_len == 0 || old_len >= BUF_CHARS)
        old_path[0] = L'\0';

    if (path_is_dir(lib_dir)) {
        StringCchPrintfW(new_path, _countof(new_path), L"%s;%s;%s",
                         g_runtime, lib_dir, old_path);
    } else {
        StringCchPrintfW(new_path, _countof(new_path), L"%s;%s",
                         g_runtime, old_path);
    }
    SetEnvironmentVariableW(L"PATH", new_path);
}

int wmain(void)
{
    const wchar_t *rest;
    STARTUPINFOW si;
    PROCESS_INFORMATION pi;
    DWORD exit_code = 1;
    DWORD err;
    wchar_t err_text[512];

    if (!resolve_root()) {
        out_text(L"\n[错误] 无法确定启动器所在目录。\n");
        press_any_key();
        return 2;
    }

    if (FAILED(join_path(g_runtime, _countof(g_runtime), g_root, RUNTIME_DIR_NAME)) ||
        FAILED(join_path(g_main_exe, _countof(g_main_exe), g_runtime, MAIN_EXE_NAME))) {
        out_text(L"\n[错误] 路径过长，请把工具放到更短、没有特殊字符的目录下。\n");
        press_any_key();
        return 2;
    }

    if (!path_is_dir(g_runtime)) {
        out_fmt(L"\n[错误] 找不到运行时目录：%s\n", g_runtime);
        out_text(L"请确认 _runtime 文件夹没有被删除或改名，或重新解压本工具。\n");
        press_any_key();
        return 3;
    }

    if (!path_is_file(g_main_exe)) {
        out_fmt(L"\n[错误] 找不到主程序：%s\n", g_main_exe);
        out_text(L"请确认 _runtime 文件夹没有被删除或改名，或重新解压本工具。\n");
        press_any_key();
        return 3;
    }

    ensure_runtime_hidden();
    prepare_environment();

    rest = args_after_exe();
    if (FAILED(StringCchPrintfW(g_cmdline, _countof(g_cmdline), L"\"%s\"%s%s",
                                g_main_exe, rest[0] != L'\0' ? L" " : L"", rest))) {
        out_text(L"\n[错误] 命令行参数过长。\n");
        press_any_key();
        return 2;
    }

    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));

    SetConsoleTitleW(L"客户Excel处理工具");

    /* 工作目录固定在根目录，程序里的相对路径才不会跑到 _runtime 里 */
    if (!CreateProcessW(g_main_exe, g_cmdline, NULL, NULL, TRUE, 0, NULL,
                        g_root, &si, &pi)) {
        err = GetLastError();
        err_text[0] = L'\0';
        FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                       NULL, err, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                       err_text, _countof(err_text), NULL);
        out_fmt(L"\n[错误] 启动主程序失败（Win32 错误 %lu）：%s\n", err, err_text);
        out_text(L"常见原因：_runtime 里的 DLL 被安全软件隔离，或文件没解压完整。\n");
        press_any_key();
        return 4;
    }

    CloseHandle(pi.hThread);
    WaitForSingleObject(pi.hProcess, INFINITE);
    GetExitCodeProcess(pi.hProcess, &exit_code);
    CloseHandle(pi.hProcess);

    return (int)exit_code;
}
