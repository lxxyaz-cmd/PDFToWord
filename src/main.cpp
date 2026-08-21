#define UNICODE
#define _UNICODE
#include <windows.h>
#include <commdlg.h>
#include <commctrl.h>
#include <shellapi.h>
#include <shlobj.h>
#include <objbase.h>
#include <oleauto.h>

#include <algorithm>
#include <array>
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <memory>
#include <sstream>
#include <string>
#include <vector>

#include "embedded_tools.h"

namespace fs = std::filesystem;

namespace {

constexpr wchar_t kClassName[] = L"OfflinePdfToWordWindow";
constexpr UINT WM_CONVERSION_DONE = WM_APP + 1;
constexpr UINT WM_STATUS_TEXT = WM_APP + 2;

enum ControlId {
    IDC_INPUT = 101,
    IDC_INPUT_BROWSE,
    IDC_OUTPUT,
    IDC_OUTPUT_BROWSE,
    IDC_ENGINE,
    IDC_CONVERT,
    IDC_OPEN_FOLDER,
    IDC_STATUS,
    IDC_APP_ICON,
    IDC_TITLE,
    IDC_SUBTITLE,
    IDC_STATUS_LABEL
};

struct AppState {
    HWND window{};
    HWND input{};
    HWND output{};
    HWND engine{};
    HWND convert{};
    HWND openFolder{};
    HWND status{};
    HWND appIcon{};
    HFONT font{};
    HFONT titleFont{};
    HBRUSH statusBrush{};
    bool busy{};
};

struct ConversionJob {
    HWND window{};
    std::wstring input;
    std::wstring output;
    int engine{}; // 0 auto, 1 editable offline, 2 page-image, 3 Word COM
};

struct ZipEntry {
    std::string name;
    std::vector<unsigned char> data;
    uint32_t crc{};
    uint32_t offset{};
};

std::wstring g_cliInput;
std::wstring g_cliOutput;
int g_cliEngine = 1;
bool g_cliMode = false;
std::wstring g_screenshotPath;
fs::path g_embeddedToolsDirectory;

std::wstring LastErrorText(DWORD error = GetLastError()) {
    wchar_t* buffer = nullptr;
    const DWORD size = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER |
                                          FORMAT_MESSAGE_FROM_SYSTEM |
                                          FORMAT_MESSAGE_IGNORE_INSERTS,
                                      nullptr, error, 0,
                                      reinterpret_cast<wchar_t*>(&buffer), 0, nullptr);
    std::wstring result = size && buffer ? std::wstring(buffer, size) : L"未知错误";
    if (buffer) LocalFree(buffer);
    while (!result.empty() && (result.back() == L'\r' || result.back() == L'\n')) result.pop_back();
    return result;
}

std::wstring HResultText(HRESULT hr, const EXCEPINFO* ex = nullptr) {
    if (ex && ex->bstrDescription && SysStringLen(ex->bstrDescription)) {
        return std::wstring(ex->bstrDescription, SysStringLen(ex->bstrDescription));
    }
    return LastErrorText(static_cast<DWORD>(hr));
}

std::wstring GetText(HWND control) {
    const int length = GetWindowTextLengthW(control);
    std::wstring text(length, L'\0');
    if (length) GetWindowTextW(control, text.data(), length + 1);
    return text;
}

void SetStatus(HWND window, const std::wstring& text) {
    if (!window) return;
    auto* copy = new std::wstring(text);
    PostMessageW(window, WM_STATUS_TEXT, 0, reinterpret_cast<LPARAM>(copy));
}

std::wstring ExecutableDirectory() {
    std::vector<wchar_t> path(32768);
    const DWORD len = GetModuleFileNameW(nullptr, path.data(), static_cast<DWORD>(path.size()));
    return fs::path(std::wstring(path.data(), len)).parent_path().wstring();
}

void CleanupEmbeddedTools() {
    if (g_embeddedToolsDirectory.empty()) return;
    std::error_code ec;
    fs::remove_all(g_embeddedToolsDirectory, ec);
    g_embeddedToolsDirectory.clear();
}

bool EnsureEmbeddedTools(HWND notifyWindow, fs::path& directory, std::wstring& error) {
    if (!g_embeddedToolsDirectory.empty() && fs::exists(g_embeddedToolsDirectory)) {
        directory = g_embeddedToolsDirectory;
        return true;
    }

    wchar_t tempRoot[MAX_PATH]{};
    if (!GetTempPathW(MAX_PATH, tempRoot)) {
        error = L"无法取得系统临时目录：" + LastErrorText();
        return false;
    }
    const fs::path target = fs::path(tempRoot) /
        (L"PDFToWord_Engine_" + std::to_wstring(GetCurrentProcessId()) + L"_" +
         std::to_wstring(GetTickCount64()));
    std::error_code ec;
    fs::create_directories(target, ec);
    if (ec) {
        error = L"无法创建单文件引擎临时目录。";
        return false;
    }

    SetStatus(notifyWindow, L"正在释放单文件内置转换引擎…");
    for (const auto& item : kEmbeddedToolResources) {
        HRSRC resource = FindResourceW(nullptr, MAKEINTRESOURCEW(item.id), RT_RCDATA);
        if (!resource) {
            error = L"单文件内置资源缺失：" + std::wstring(item.name);
            fs::remove_all(target, ec);
            return false;
        }
        HGLOBAL loaded = LoadResource(nullptr, resource);
        const DWORD size = SizeofResource(nullptr, resource);
        const void* bytes = loaded ? LockResource(loaded) : nullptr;
        if (!bytes || size == 0) {
            error = L"无法读取单文件内置资源：" + std::wstring(item.name);
            fs::remove_all(target, ec);
            return false;
        }
        const fs::path output = target / item.name;
        std::ofstream stream(output, std::ios::binary | std::ios::trunc);
        stream.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(size));
        if (!stream.good()) {
            error = L"无法释放单文件内置资源：" + output.wstring();
            fs::remove_all(target, ec);
            return false;
        }
    }
    g_embeddedToolsDirectory = target;
    directory = target;
    return true;
}

std::wstring QuoteArgument(const std::wstring& value) {
    std::wstring quoted = L"\"";
    size_t slashes = 0;
    for (wchar_t ch : value) {
        if (ch == L'\\') {
            ++slashes;
        } else if (ch == L'\"') {
            quoted.append(slashes * 2 + 1, L'\\');
            quoted.push_back(L'\"');
            slashes = 0;
        } else {
            quoted.append(slashes, L'\\');
            slashes = 0;
            quoted.push_back(ch);
        }
    }
    quoted.append(slashes * 2, L'\\');
    quoted.push_back(L'\"');
    return quoted;
}

bool RunProcess(const std::wstring& executable, const std::wstring& arguments,
                DWORD& exitCode, std::wstring& error) {
    std::wstring command = QuoteArgument(executable) + L" " + arguments;
    STARTUPINFOW startup{};
    startup.cb = sizeof(startup);
    startup.dwFlags = STARTF_USESHOWWINDOW;
    startup.wShowWindow = SW_HIDE;
    PROCESS_INFORMATION process{};
    if (!CreateProcessW(executable.c_str(), command.data(), nullptr, nullptr, FALSE,
                        CREATE_NO_WINDOW, nullptr, nullptr, &startup, &process)) {
        error = L"无法启动离线 PDF 引擎：" + LastErrorText();
        return false;
    }
    WaitForSingleObject(process.hProcess, INFINITE);
    GetExitCodeProcess(process.hProcess, &exitCode);
    CloseHandle(process.hThread);
    CloseHandle(process.hProcess);
    if (exitCode != 0) {
        error = L"离线 PDF 引擎转换失败，退出代码：" + std::to_wstring(exitCode);
        return false;
    }
    return true;
}

uint32_t Crc32(const std::vector<unsigned char>& data) {
    static std::array<uint32_t, 256> table{};
    static bool initialized = false;
    if (!initialized) {
        for (uint32_t i = 0; i < 256; ++i) {
            uint32_t c = i;
            for (int j = 0; j < 8; ++j) c = (c & 1) ? (0xedb88320U ^ (c >> 1)) : (c >> 1);
            table[i] = c;
        }
        initialized = true;
    }
    uint32_t crc = 0xffffffffU;
    for (unsigned char byte : data) crc = table[(crc ^ byte) & 0xff] ^ (crc >> 8);
    return crc ^ 0xffffffffU;
}

void WriteU16(std::ofstream& stream, uint16_t value) {
    stream.put(static_cast<char>(value & 0xff));
    stream.put(static_cast<char>((value >> 8) & 0xff));
}

void WriteU32(std::ofstream& stream, uint32_t value) {
    WriteU16(stream, static_cast<uint16_t>(value & 0xffff));
    WriteU16(stream, static_cast<uint16_t>((value >> 16) & 0xffff));
}

std::vector<unsigned char> Bytes(const std::string& text) {
    return std::vector<unsigned char>(text.begin(), text.end());
}

bool ReadFileBytes(const fs::path& path, std::vector<unsigned char>& data) {
    std::ifstream stream(path, std::ios::binary);
    if (!stream) return false;
    stream.seekg(0, std::ios::end);
    const auto size = stream.tellg();
    stream.seekg(0, std::ios::beg);
    if (size < 0) return false;
    data.resize(static_cast<size_t>(size));
    if (!data.empty()) stream.read(reinterpret_cast<char*>(data.data()), size);
    return stream.good() || stream.eof();
}

std::wstring ReadUtf8Text(const fs::path& path) {
    std::vector<unsigned char> data;
    if (!ReadFileBytes(path, data) || data.empty()) return {};
    const int length = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS,
                                            reinterpret_cast<const char*>(data.data()),
                                            static_cast<int>(data.size()), nullptr, 0);
    if (length <= 0) return {};
    std::wstring result(static_cast<size_t>(length), L'\0');
    MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS,
                        reinterpret_cast<const char*>(data.data()),
                        static_cast<int>(data.size()), result.data(), length);
    return result;
}

bool WriteStoredZip(const fs::path& output, std::vector<ZipEntry>& entries, std::wstring& error) {
    std::ofstream stream(output, std::ios::binary | std::ios::trunc);
    if (!stream) {
        error = L"无法创建输出文件，请检查保存路径和权限。";
        return false;
    }
    for (auto& entry : entries) {
        entry.crc = Crc32(entry.data);
        entry.offset = static_cast<uint32_t>(stream.tellp());
        WriteU32(stream, 0x04034b50);
        WriteU16(stream, 20);
        WriteU16(stream, 0x0800); // UTF-8 names
        WriteU16(stream, 0);      // stored
        WriteU16(stream, 0);
        WriteU16(stream, 0);
        WriteU32(stream, entry.crc);
        WriteU32(stream, static_cast<uint32_t>(entry.data.size()));
        WriteU32(stream, static_cast<uint32_t>(entry.data.size()));
        WriteU16(stream, static_cast<uint16_t>(entry.name.size()));
        WriteU16(stream, 0);
        stream.write(entry.name.data(), entry.name.size());
        if (!entry.data.empty()) stream.write(reinterpret_cast<const char*>(entry.data.data()), entry.data.size());
    }
    const uint32_t centralOffset = static_cast<uint32_t>(stream.tellp());
    for (const auto& entry : entries) {
        WriteU32(stream, 0x02014b50);
        WriteU16(stream, 20);
        WriteU16(stream, 20);
        WriteU16(stream, 0x0800);
        WriteU16(stream, 0);
        WriteU16(stream, 0);
        WriteU16(stream, 0);
        WriteU32(stream, entry.crc);
        WriteU32(stream, static_cast<uint32_t>(entry.data.size()));
        WriteU32(stream, static_cast<uint32_t>(entry.data.size()));
        WriteU16(stream, static_cast<uint16_t>(entry.name.size()));
        WriteU16(stream, 0);
        WriteU16(stream, 0);
        WriteU16(stream, 0);
        WriteU16(stream, 0);
        WriteU32(stream, 0);
        WriteU32(stream, entry.offset);
        stream.write(entry.name.data(), entry.name.size());
    }
    const uint32_t centralSize = static_cast<uint32_t>(stream.tellp()) - centralOffset;
    WriteU32(stream, 0x06054b50);
    WriteU16(stream, 0);
    WriteU16(stream, 0);
    WriteU16(stream, static_cast<uint16_t>(entries.size()));
    WriteU16(stream, static_cast<uint16_t>(entries.size()));
    WriteU32(stream, centralSize);
    WriteU32(stream, centralOffset);
    WriteU16(stream, 0);
    if (!stream.good()) {
        error = L"写入 DOCX 时发生磁盘错误。";
        return false;
    }
    return true;
}

bool PngDimensions(const std::vector<unsigned char>& data, uint32_t& width, uint32_t& height) {
    if (data.size() < 24 || data[0] != 0x89 || data[1] != 'P' || data[2] != 'N' || data[3] != 'G') return false;
    auto be32 = [&](size_t offset) {
        return (static_cast<uint32_t>(data[offset]) << 24) |
               (static_cast<uint32_t>(data[offset + 1]) << 16) |
               (static_cast<uint32_t>(data[offset + 2]) << 8) |
               static_cast<uint32_t>(data[offset + 3]);
    };
    width = be32(16);
    height = be32(20);
    return width > 0 && height > 0;
}

bool BuildImageDocx(const std::vector<fs::path>& pages, const fs::path& output, std::wstring& error) {
    if (pages.empty()) {
        error = L"PDF 没有可转换的页面。";
        return false;
    }

    struct RenderedPage {
        std::vector<unsigned char> image;
        uint32_t width{};
        uint32_t height{};
    };
    std::vector<RenderedPage> renderedPages;
    renderedPages.reserve(pages.size());
    for (size_t i = 0; i < pages.size(); ++i) {
        RenderedPage page;
        if (!ReadFileBytes(pages[i], page.image) || !PngDimensions(page.image, page.width, page.height)) {
            error = L"无法读取渲染后的第 " + std::to_wstring(i + 1) + L" 页。";
            return false;
        }
        renderedPages.push_back(std::move(page));
    }

    std::vector<ZipEntry> entries;
    std::ostringstream document;
    std::ostringstream relationships;
    document << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
             << "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
             << "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
             << "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
             << "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
             << "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><w:body>";
    relationships << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                  << "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";

    for (size_t i = 0; i < renderedPages.size(); ++i) {
        auto& page = renderedPages[i];
        const uint32_t width = page.width;
        const uint32_t height = page.height;
        // pdftoppm renders at 144 DPI: 1 pixel = 6,350 EMU.
        const uint64_t cx = static_cast<uint64_t>(width) * 6350ULL;
        const uint64_t cy = static_cast<uint64_t>(height) * 6350ULL;
        const size_t id = i + 1;
        document << "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/>"
                 << "<w:jc w:val=\"center\"/>";
        if (i + 1 < renderedPages.size()) {
            document << "<w:sectPr><w:type w:val=\"nextPage\"/><w:pgSz w:w=\"" << width * 10
                     << "\" w:h=\"" << height * 10 << "\"";
            if (width > height) document << " w:orient=\"landscape\"";
            document << "/><w:pgMar w:top=\"0\" w:right=\"0\" w:bottom=\"0\" w:left=\"0\" w:header=\"0\" w:footer=\"0\" w:gutter=\"0\"/>"
                     << "</w:sectPr>";
        }
        document << "</w:pPr><w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">"
                 << "<wp:extent cx=\"" << cx << "\" cy=\"" << cy << "\"/><wp:docPr id=\"" << id
                 << "\" name=\"PDF Page " << id << "\"/><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"
                 << "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"" << id << "\" name=\"page" << id
                 << ".png\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"rId" << id
                 << "\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\""
                 << cx << "\" cy=\"" << cy << "\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>"
                 << "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
        document << "</w:p>";
        relationships << "<Relationship Id=\"rId" << id
                      << "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/page"
                      << id << ".png\"/>";
        entries.push_back({"word/media/page" + std::to_string(id) + ".png", std::move(page.image)});
    }
    const auto& lastPage = renderedPages.back();
    const uint32_t pageWidthTwips = lastPage.width * 10;
    const uint32_t pageHeightTwips = lastPage.height * 10;
    document << "<w:sectPr><w:pgSz w:w=\"" << pageWidthTwips << "\" w:h=\"" << pageHeightTwips
             << "\"";
    if (lastPage.width > lastPage.height) document << " w:orient=\"landscape\"";
    document << "/><w:pgMar w:top=\"0\" w:right=\"0\" w:bottom=\"0\" w:left=\"0\" w:header=\"0\" w:footer=\"0\" w:gutter=\"0\"/>"
             << "</w:sectPr></w:body></w:document>";
    relationships << "</Relationships>";

    const std::string contentTypes =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Default Extension=\"png\" ContentType=\"image/png\"/>"
        "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
        "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>"
        "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>"
        "</Types>";
    const std::string rootRels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>"
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"
        "</Relationships>";
    const std::string core =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
        "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" "
        "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:title>PDF conversion</dc:title>"
        "<dc:creator>PDF to Word Offline</dc:creator></cp:coreProperties>";
    const std::string app =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" "
        "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">"
        "<Application>PDF to Word Offline</Application><AppVersion>1.0</AppVersion></Properties>";

    entries.insert(entries.begin(), {
        {"[Content_Types].xml", Bytes(contentTypes)},
        {"_rels/.rels", Bytes(rootRels)},
        {"word/document.xml", Bytes(document.str())},
        {"word/_rels/document.xml.rels", Bytes(relationships.str())},
        {"docProps/core.xml", Bytes(core)},
        {"docProps/app.xml", Bytes(app)}
    });
    return WriteStoredZip(output, entries, error);
}

bool ConvertWithImages(const std::wstring& input, const std::wstring& output,
                       HWND notifyWindow, std::wstring& error) {
    fs::path tools;
    if (!EnsureEmbeddedTools(notifyWindow, tools, error)) return false;
    const fs::path tool = tools / L"pdftoppm.exe";
    wchar_t tempRoot[MAX_PATH]{};
    if (!GetTempPathW(MAX_PATH, tempRoot)) {
        error = L"无法取得系统临时目录：" + LastErrorText();
        return false;
    }
    const fs::path temp = fs::path(tempRoot) /
        (L"PDFToWord_" + std::to_wstring(GetCurrentProcessId()) + L"_" + std::to_wstring(GetTickCount64()));
    std::error_code ec;
    fs::create_directories(temp, ec);
    if (ec) {
        error = L"无法创建临时目录。";
        return false;
    }
    const fs::path prefix = temp / L"page";
    SetStatus(notifyWindow, L"正在离线渲染 PDF 页面…");
    DWORD exitCode = 0;
    const std::wstring arguments = L"-png -r 144 " + QuoteArgument(input) + L" " + QuoteArgument(prefix.wstring());
    bool ok = RunProcess(tool.wstring(), arguments, exitCode, error);
    std::vector<fs::path> pages;
    if (ok) {
        for (const auto& item : fs::directory_iterator(temp, ec)) {
            if (item.is_regular_file() && item.path().extension() == L".png") pages.push_back(item.path());
        }
        auto pageNumber = [](const fs::path& path) {
            const std::wstring stem = path.stem().wstring();
            const auto dash = stem.find_last_of(L'-');
            return dash == std::wstring::npos ? 0 : _wtoi(stem.c_str() + dash + 1);
        };
        std::sort(pages.begin(), pages.end(), [&](const fs::path& a, const fs::path& b) { return pageNumber(a) < pageNumber(b); });
        SetStatus(notifyWindow, L"正在生成标准 DOCX 文件…");
        ok = BuildImageDocx(pages, output, error);
    }
    fs::remove_all(temp, ec);
    return ok;
}

bool ConvertEditableOffline(const std::wstring& input, const std::wstring& output,
                            HWND notifyWindow, std::wstring& error, std::wstring& warning) {
    fs::path tools;
    if (!EnsureEmbeddedTools(notifyWindow, tools, error)) return false;
    const fs::path tool = tools / L"editable_engine.exe";
    SetStatus(notifyWindow, L"正在重建版式、识别表格并校验数字完整性…");
    DWORD exitCode = 0;
    const std::wstring suffix = std::to_wstring(GetCurrentProcessId()) + L".txt";
    const fs::path errorFile = tools / (L"conversion-error-" + suffix);
    // The engine reports recoverable issues here, for example a few numbers
    // that only exist inside a drawing. Those no longer fail the conversion,
    // so the finished document is delivered together with the remark.
    const fs::path warningFile = tools / (L"conversion-warning-" + suffix);
    const std::wstring arguments = QuoteArgument(input) + L" " + QuoteArgument(output) +
        L" --error-file " + QuoteArgument(errorFile.wstring()) +
        L" --warning-file " + QuoteArgument(warningFile.wstring());
    std::error_code cleanupError;
    if (!RunProcess(tool.wstring(), arguments, exitCode, error)) {
        const std::wstring detail = ReadUtf8Text(errorFile);
        fs::remove(errorFile, cleanupError);
        fs::remove(warningFile, cleanupError);
        error = detail.empty()
            ? L"智能可编辑转换失败：PDF 可能已损坏、加密、数字校验未通过，或页面结构过于特殊。"
            : L"智能可编辑转换失败：" + detail;
        return false;
    }
    warning = ReadUtf8Text(warningFile);
    fs::remove(errorFile, cleanupError);
    fs::remove(warningFile, cleanupError);
    if (!fs::exists(output) || fs::file_size(output) == 0) {
        error = L"可编辑离线引擎没有生成有效的 DOCX 文件。";
        return false;
    }
    return true;
}

HRESULT GetDispatchProperty(IDispatch* object, const wchar_t* name, IDispatch** result, EXCEPINFO& ex) {
    DISPID id;
    LPOLESTR mutableName = const_cast<LPOLESTR>(name);
    HRESULT hr = object->GetIDsOfNames(IID_NULL, &mutableName, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) return hr;
    DISPPARAMS params{};
    VARIANT value;
    VariantInit(&value);
    UINT argError = 0;
    hr = object->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                        &params, &value, &ex, &argError);
    if (SUCCEEDED(hr) && value.vt == VT_DISPATCH) {
        *result = value.pdispVal;
    } else {
        VariantClear(&value);
        if (SUCCEEDED(hr)) hr = E_NOINTERFACE;
    }
    return hr;
}

HRESULT GetStringProperty(IDispatch* object, const wchar_t* name, std::wstring& result, EXCEPINFO& ex) {
    DISPID id;
    LPOLESTR mutableName = const_cast<LPOLESTR>(name);
    HRESULT hr = object->GetIDsOfNames(IID_NULL, &mutableName, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) return hr;
    DISPPARAMS params{};
    VARIANT value;
    VariantInit(&value);
    UINT argError = 0;
    hr = object->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                        &params, &value, &ex, &argError);
    if (SUCCEEDED(hr) && value.vt == VT_BSTR && value.bstrVal) {
        result.assign(value.bstrVal, SysStringLen(value.bstrVal));
    } else if (SUCCEEDED(hr)) {
        hr = DISP_E_TYPEMISMATCH;
    }
    VariantClear(&value);
    return hr;
}

HRESULT PutIntProperty(IDispatch* object, const wchar_t* name, long value, EXCEPINFO& ex) {
    DISPID id;
    LPOLESTR mutableName = const_cast<LPOLESTR>(name);
    HRESULT hr = object->GetIDsOfNames(IID_NULL, &mutableName, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) return hr;
    DISPID named = DISPID_PROPERTYPUT;
    VARIANT argument;
    VariantInit(&argument);
    argument.vt = VT_I4;
    argument.lVal = value;
    DISPPARAMS params{&argument, &named, 1, 1};
    UINT argError = 0;
    return object->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                          &params, nullptr, &ex, &argError);
}

HRESULT CallMethod(IDispatch* object, const wchar_t* name, std::vector<VARIANT>& arguments,
                   VARIANT* result, EXCEPINFO& ex) {
    DISPID id;
    LPOLESTR mutableName = const_cast<LPOLESTR>(name);
    HRESULT hr = object->GetIDsOfNames(IID_NULL, &mutableName, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) return hr;
    DISPPARAMS params{};
    params.rgvarg = arguments.empty() ? nullptr : arguments.data();
    params.cArgs = static_cast<UINT>(arguments.size());
    UINT argError = 0;
    return object->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &params, result, &ex, &argError);
}

VARIANT StringVariant(const std::wstring& value) {
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_BSTR;
    v.bstrVal = SysAllocStringLen(value.data(), static_cast<UINT>(value.size()));
    return v;
}

VARIANT IntVariant(long value) {
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_I4;
    v.lVal = value;
    return v;
}

bool WordAvailable() {
    CLSID clsid{};
    return SUCCEEDED(CLSIDFromProgID(L"Word.Application", &clsid));
}

bool WordOpenedPdfAsRawText(IDispatch* document) {
    IDispatch* content = nullptr;
    EXCEPINFO ex{};
    if (FAILED(GetDispatchProperty(document, L"Content", &content, ex)) || !content) return false;
    std::wstring text;
    const HRESULT hr = GetStringProperty(content, L"Text", text, ex);
    content->Release();
    if (FAILED(hr)) return false;

    const size_t sampleLength = (std::min)(text.size(), static_cast<size_t>(8192));
    const std::wstring sample = text.substr(0, sampleLength);
    const bool hasPdfHeader = sample.find(L"%PDF-") != std::wstring::npos;
    const bool hasPdfObjects = sample.find(L" obj") != std::wstring::npos &&
                               sample.find(L"stream") != std::wstring::npos;
    return hasPdfHeader || hasPdfObjects;
}

bool ConvertWithWord(const std::wstring& input, const std::wstring& output, std::wstring& error) {
    const HRESULT init = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    const bool uninitialize = SUCCEEDED(init);
    CLSID clsid{};
    HRESULT hr = CLSIDFromProgID(L"Word.Application", &clsid);
    IDispatch* word = nullptr;
    IDispatch* documents = nullptr;
    IDispatch* document = nullptr;
    EXCEPINFO ex{};

    if (SUCCEEDED(hr)) hr = CoCreateInstance(clsid, nullptr, CLSCTX_LOCAL_SERVER, IID_IDispatch, reinterpret_cast<void**>(&word));
    if (FAILED(hr)) {
        error = L"未检测到可用的 Microsoft Word。请选择“精准版式”模式，或安装 Word 后重试。";
        if (uninitialize) CoUninitialize();
        return false;
    }
    PutIntProperty(word, L"Visible", 0, ex);
    PutIntProperty(word, L"DisplayAlerts", 0, ex);
    hr = GetDispatchProperty(word, L"Documents", &documents, ex);
    if (SUCCEEDED(hr)) {
        std::vector<VARIANT> args;
        args.push_back(StringVariant(input));
        VARIANT result;
        VariantInit(&result);
        hr = CallMethod(documents, L"Open", args, &result, ex);
        VariantClear(&args[0]);
        if (SUCCEEDED(hr) && result.vt == VT_DISPATCH) document = result.pdispVal;
        else VariantClear(&result);
    }
    if (SUCCEEDED(hr) && document) {
        if (WordOpenedPdfAsRawText(document)) {
            hr = HRESULT_FROM_WIN32(ERROR_BAD_FORMAT);
            error = L"Microsoft Word 未启用或不支持 PDF Reflow，已将 PDF 误识别为原始文本。请使用智能可编辑模式。";
        } else {
            std::vector<VARIANT> args;
            args.push_back(IntVariant(16)); // wdFormatDocumentDefault (.docx)
            args.push_back(StringVariant(output));
            hr = CallMethod(document, L"SaveAs2", args, nullptr, ex);
            VariantClear(&args[1]);
        }
    }
    if (document) {
        std::vector<VARIANT> args{IntVariant(0)};
        EXCEPINFO closeEx{};
        CallMethod(document, L"Close", args, nullptr, closeEx);
        document->Release();
    }
    if (word) {
        std::vector<VARIANT> noArgs;
        EXCEPINFO quitEx{};
        CallMethod(word, L"Quit", noArgs, nullptr, quitEx);
    }
    if (documents) documents->Release();
    if (word) word->Release();
    if (uninitialize) CoUninitialize();
    if (FAILED(hr)) {
        if (error.empty()) error = L"Microsoft Word 转换失败：" + HResultText(hr, &ex);
        return false;
    }
    return fs::exists(output) && fs::file_size(output) > 0;
}

bool ValidatePaths(const std::wstring& input, const std::wstring& output, std::wstring& error) {
    if (input.empty() || output.empty()) {
        error = L"请选择 PDF 文件和 Word 保存位置。";
        return false;
    }
    const fs::path source(input);
    const fs::path target(output);
    if (!fs::exists(source) || !fs::is_regular_file(source)) {
        error = L"PDF 文件不存在。";
        return false;
    }
    std::wstring extension = source.extension().wstring();
    std::transform(extension.begin(), extension.end(), extension.begin(), ::towlower);
    if (extension != L".pdf") {
        error = L"输入文件必须是 PDF。";
        return false;
    }
    const std::wstring sourceFull = fs::absolute(source).lexically_normal().wstring();
    const std::wstring targetFull = fs::absolute(target).lexically_normal().wstring();
    if (_wcsicmp(sourceFull.c_str(), targetFull.c_str()) == 0) {
        error = L"输入和输出不能是同一个文件。";
        return false;
    }
    std::error_code ec;
    if (!target.parent_path().empty()) fs::create_directories(target.parent_path(), ec);
    if (ec) {
        error = L"无法创建输出目录。";
        return false;
    }
    return true;
}

bool PerformConversion(const ConversionJob& job, std::wstring& message) {
    std::wstring error;
    if (!ValidatePaths(job.input, job.output, error)) {
        message = error;
        return false;
    }
    std::error_code ec;
    if (fs::exists(job.output)) fs::remove(job.output, ec);
    bool ok = false;
    std::wstring warning;
    if (job.engine == 3) {
        SetStatus(job.window, L"正在调用 Microsoft Word 转换（请勿关闭 Word）…");
        ok = ConvertWithWord(job.input, job.output, error);
    } else if (job.engine == 2) {
        ok = ConvertWithImages(job.input, job.output, job.window, error);
    } else {
        // Auto deliberately uses the verified offline engine. Word registration
        // does not guarantee that PDF Reflow is installed or enabled.
        ok = ConvertEditableOffline(job.input, job.output, job.window, error, warning);
    }
    if (ok) {
        message = L"转换完成：" + job.output;
        if (!warning.empty()) message += L"\r\n\r\n请注意：\r\n" + warning;
    } else {
        message = error.empty() ? L"转换失败。" : error;
    }
    return ok;
}

DWORD WINAPI ConversionThread(void* parameter) {
    std::unique_ptr<ConversionJob> job(static_cast<ConversionJob*>(parameter));
    std::wstring message;
    bool ok = false;
    try {
        ok = PerformConversion(*job, message);
    } catch (const std::exception&) {
        message = L"转换过程中发生文件或内存错误。请检查 PDF 是否损坏、磁盘空间是否充足。";
    } catch (...) {
        message = L"转换过程中发生未知错误。";
    }
    auto* copy = new std::wstring(std::move(message));
    PostMessageW(job->window, WM_CONVERSION_DONE, ok ? 1 : 0, reinterpret_cast<LPARAM>(copy));
    return 0;
}

void ChooseInput(AppState* state) {
    wchar_t path[32768]{};
    OPENFILENAMEW dialog{};
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = state->window;
    dialog.lpstrFilter = L"PDF 文件 (*.pdf)\0*.pdf\0所有文件 (*.*)\0*.*\0";
    dialog.lpstrFile = path;
    dialog.nMaxFile = static_cast<DWORD>(std::size(path));
    dialog.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
    dialog.lpstrDefExt = L"pdf";
    if (GetOpenFileNameW(&dialog)) {
        SetWindowTextW(state->input, path);
        fs::path output(path);
        output.replace_extension(L".docx");
        SetWindowTextW(state->output, output.c_str());
    }
}

void ChooseOutput(AppState* state) {
    wchar_t path[32768]{};
    const std::wstring current = GetText(state->output);
    wcsncpy_s(path, current.c_str(), _TRUNCATE);
    OPENFILENAMEW dialog{};
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = state->window;
    dialog.lpstrFilter = L"Word 文档 (*.docx)\0*.docx\0";
    dialog.lpstrFile = path;
    dialog.nMaxFile = static_cast<DWORD>(std::size(path));
    dialog.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;
    dialog.lpstrDefExt = L"docx";
    if (GetSaveFileNameW(&dialog)) SetWindowTextW(state->output, path);
}

void StartConversion(AppState* state) {
    const std::wstring input = GetText(state->input);
    const std::wstring output = GetText(state->output);
    std::wstring error;
    if (!ValidatePaths(input, output, error)) {
        MessageBoxW(state->window, error.c_str(), L"无法开始转换", MB_OK | MB_ICONWARNING);
        return;
    }
    if (fs::exists(output) && MessageBoxW(state->window, L"输出文件已存在，是否覆盖？", L"确认覆盖",
                                           MB_YESNO | MB_ICONQUESTION) != IDYES) return;
    auto* job = new ConversionJob{state->window, input, output,
                                  static_cast<int>(SendMessageW(state->engine, CB_GETCURSEL, 0, 0))};
    state->busy = true;
    EnableWindow(state->convert, FALSE);
    EnableWindow(state->input, FALSE);
    EnableWindow(state->output, FALSE);
    SetWindowTextW(state->status, L"正在准备转换…");
    HANDLE thread = CreateThread(nullptr, 0, ConversionThread, job, 0, nullptr);
    if (!thread) {
        delete job;
        state->busy = false;
        EnableWindow(state->convert, TRUE);
        EnableWindow(state->input, TRUE);
        EnableWindow(state->output, TRUE);
        MessageBoxW(state->window, LastErrorText().c_str(), L"启动失败", MB_OK | MB_ICONERROR);
        return;
    }
    CloseHandle(thread);
}

void LayoutControls(AppState* state, int width, int height) {
    const int margin = 48;
    const int buttonWidth = 120;
    const int editWidth = width - margin * 2 - buttonWidth - 12;
    MoveWindow(state->appIcon, margin, 38, 64, 64, TRUE);
    MoveWindow(GetDlgItem(state->window, IDC_TITLE), 128, 34, width - 176, 40, TRUE);
    MoveWindow(GetDlgItem(state->window, IDC_SUBTITLE), 130, 76, width - 178, 24, TRUE);
    MoveWindow(state->input, margin, 176, editWidth, 40, TRUE);
    MoveWindow(GetDlgItem(state->window, IDC_INPUT_BROWSE), margin + editWidth + 12, 176, buttonWidth, 40, TRUE);
    MoveWindow(state->output, margin, 266, editWidth, 40, TRUE);
    MoveWindow(GetDlgItem(state->window, IDC_OUTPUT_BROWSE), margin + editWidth + 12, 266, buttonWidth, 40, TRUE);
    MoveWindow(state->engine, margin, 356, width - margin * 2, 42, TRUE);
    MoveWindow(state->convert, margin, 424, 196, 50, TRUE);
    MoveWindow(state->openFolder, margin + 212, 424, 148, 50, TRUE);
    MoveWindow(GetDlgItem(state->window, IDC_STATUS_LABEL), margin, 498, 160, 24, TRUE);
    const int statusHeight = height - 562 > 62 ? height - 562 : 62;
    MoveWindow(state->status, margin, 530, width - margin * 2, statusHeight, TRUE);
}

HWND AddStatic(HWND parent, const wchar_t* text, int x, int y, int width, int height, HFONT font, int id = 0) {
    HWND label = CreateWindowExW(0, L"STATIC", text, WS_CHILD | WS_VISIBLE,
                                 x, y, width, height, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), GetModuleHandleW(nullptr), nullptr);
    SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    return label;
}

void DrawBackground(HWND window, HDC dc) {
    RECT client{};
    GetClientRect(window, &client);
    HBRUSH outer = CreateSolidBrush(RGB(241, 245, 252));
    FillRect(dc, &client, outer);
    DeleteObject(outer);

    RECT shadow{23, 20, client.right - 17, client.bottom - 13};
    HBRUSH shadowBrush = CreateSolidBrush(RGB(218, 226, 240));
    HGDIOBJ oldShadowBrush = SelectObject(dc, shadowBrush);
    HGDIOBJ oldShadowPen = SelectObject(dc, GetStockObject(NULL_PEN));
    RoundRect(dc, shadow.left, shadow.top, shadow.right, shadow.bottom, 24, 24);
    SelectObject(dc, oldShadowPen);
    SelectObject(dc, oldShadowBrush);
    DeleteObject(shadowBrush);

    RECT card{18, 14, client.right - 22, client.bottom - 18};
    HBRUSH cardBrush = CreateSolidBrush(RGB(255, 255, 255));
    HPEN border = CreatePen(PS_SOLID, 1, RGB(211, 223, 244));
    HGDIOBJ oldBrush = SelectObject(dc, cardBrush);
    HGDIOBJ oldPen = SelectObject(dc, border);
    RoundRect(dc, card.left, card.top, card.right, card.bottom, 24, 24);
    SelectObject(dc, oldPen);
    SelectObject(dc, oldBrush);
    DeleteObject(border);
    DeleteObject(cardBrush);

    HBRUSH headerBrush = CreateSolidBrush(RGB(10, 38, 104));
    HPEN headerPen = CreatePen(PS_SOLID, 1, RGB(10, 38, 104));
    oldBrush = SelectObject(dc, headerBrush);
    oldPen = SelectObject(dc, headerPen);
    RoundRect(dc, card.left, card.top, card.right, 126, 24, 24);
    RECT headerSquare{card.left, 92, card.right, 126};
    FillRect(dc, &headerSquare, headerBrush);
    int saved = SaveDC(dc);
    IntersectClipRect(dc, card.left, card.top, card.right, 126);
    HBRUSH glowBrush = CreateSolidBrush(RGB(20, 91, 214));
    HGDIOBJ oldGlow = SelectObject(dc, glowBrush);
    Ellipse(dc, client.right - 230, -72, client.right + 80, 205);
    SelectObject(dc, oldGlow);
    DeleteObject(glowBrush);
    RestoreDC(dc, saved);
    SelectObject(dc, oldPen);
    SelectObject(dc, oldBrush);
    DeleteObject(headerPen);
    DeleteObject(headerBrush);

    HBRUSH cyan = CreateSolidBrush(RGB(41, 209, 255));
    RECT cyanLine{card.left, 123, card.right, 126};
    FillRect(dc, &cyanLine, cyan);
    DeleteObject(cyan);

    RECT formCard{32, 134, client.right - 36, 410};
    HBRUSH formBrush = CreateSolidBrush(RGB(248, 250, 255));
    HPEN formPen = CreatePen(PS_SOLID, 1, RGB(218, 227, 243));
    oldBrush = SelectObject(dc, formBrush);
    oldPen = SelectObject(dc, formPen);
    RoundRect(dc, formCard.left, formCard.top, formCard.right, formCard.bottom, 16, 16);
    SelectObject(dc, oldPen);
    SelectObject(dc, oldBrush);
    DeleteObject(formPen);
    DeleteObject(formBrush);
}

void DrawFolderGlyph(HDC dc, int x, int y, COLORREF color) {
    HPEN pen = CreatePen(PS_SOLID, 2, color);
    HGDIOBJ oldPen = SelectObject(dc, pen);
    HGDIOBJ oldBrush = SelectObject(dc, GetStockObject(NULL_BRUSH));
    POINT points[] = {{x, y + 5}, {x + 7, y + 5}, {x + 10, y + 1}, {x + 20, y + 1},
                      {x + 23, y + 6}, {x + 23, y + 18}, {x, y + 18}, {x, y + 5}};
    Polyline(dc, points, static_cast<int>(std::size(points)));
    SelectObject(dc, oldBrush);
    SelectObject(dc, oldPen);
    DeleteObject(pen);
}

void DrawConvertGlyph(HDC dc, int x, int y, COLORREF color) {
    HPEN pen = CreatePen(PS_SOLID, 2, color);
    HGDIOBJ old = SelectObject(dc, pen);
    MoveToEx(dc, x, y + 4, nullptr); LineTo(dc, x + 19, y + 4);
    MoveToEx(dc, x + 15, y, nullptr); LineTo(dc, x + 20, y + 4); LineTo(dc, x + 15, y + 8);
    MoveToEx(dc, x + 20, y + 15, nullptr); LineTo(dc, x + 1, y + 15);
    MoveToEx(dc, x + 5, y + 11, nullptr); LineTo(dc, x, y + 15); LineTo(dc, x + 5, y + 19);
    SelectObject(dc, old);
    DeleteObject(pen);
}

void DrawOwnerButton(const DRAWITEMSTRUCT* item) {
    const bool primary = item->CtlID == IDC_CONVERT;
    const bool pressed = (item->itemState & ODS_SELECTED) != 0;
    const bool disabled = (item->itemState & ODS_DISABLED) != 0;
    COLORREF fill = primary ? (pressed ? RGB(10, 54, 154) : RGB(16, 91, 215))
                            : (pressed ? RGB(230, 239, 255) : RGB(255, 255, 255));
    COLORREF borderColor = primary ? RGB(16, 91, 215) : RGB(182, 199, 229);
    COLORREF textColor = disabled ? RGB(148, 163, 184) : (primary ? RGB(255, 255, 255) : RGB(22, 66, 158));
    HBRUSH shadow = CreateSolidBrush(RGB(211, 222, 241));
    HGDIOBJ oldShadow = SelectObject(item->hDC, shadow);
    HGDIOBJ oldShadowPen = SelectObject(item->hDC, GetStockObject(NULL_PEN));
    RoundRect(item->hDC, item->rcItem.left, item->rcItem.top + 2,
              item->rcItem.right, item->rcItem.bottom + 2, 14, 14);
    SelectObject(item->hDC, oldShadowPen);
    SelectObject(item->hDC, oldShadow);
    DeleteObject(shadow);
    HBRUSH brush = CreateSolidBrush(fill);
    HPEN pen = CreatePen(PS_SOLID, 1, borderColor);
    HGDIOBJ oldBrush = SelectObject(item->hDC, brush);
    HGDIOBJ oldPen = SelectObject(item->hDC, pen);
    RoundRect(item->hDC, item->rcItem.left, item->rcItem.top,
              item->rcItem.right, item->rcItem.bottom, 14, 14);
    SelectObject(item->hDC, oldPen);
    SelectObject(item->hDC, oldBrush);
    DeleteObject(pen);
    DeleteObject(brush);

    wchar_t text[128]{};
    GetWindowTextW(item->hwndItem, text, static_cast<int>(std::size(text)));
    RECT textRect = item->rcItem;
    textRect.left += 30;
    SetBkMode(item->hDC, TRANSPARENT);
    SetTextColor(item->hDC, textColor);
    HFONT font = reinterpret_cast<HFONT>(SendMessageW(item->hwndItem, WM_GETFONT, 0, 0));
    HGDIOBJ oldFont = SelectObject(item->hDC, font);
    DrawTextW(item->hDC, text, -1, &textRect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    SelectObject(item->hDC, oldFont);
    const int glyphX = item->rcItem.left + 16;
    const int glyphY = item->rcItem.top + (item->rcItem.bottom - item->rcItem.top - 20) / 2;
    if (item->CtlID == IDC_CONVERT) DrawConvertGlyph(item->hDC, glyphX, glyphY, textColor);
    else DrawFolderGlyph(item->hDC, glyphX, glyphY, textColor);
}

bool CaptureWindowBmp(HWND window, const std::wstring& path) {
    RECT rect{};
    GetWindowRect(window, &rect);
    const int width = rect.right - rect.left;
    const int height = rect.bottom - rect.top;
    HDC source = GetWindowDC(window);
    HDC memory = CreateCompatibleDC(source);
    HBITMAP bitmap = CreateCompatibleBitmap(source, width, height);
    HGDIOBJ old = SelectObject(memory, bitmap);
    PrintWindow(window, memory, PW_RENDERFULLCONTENT);
    BITMAPINFOHEADER header{};
    header.biSize = sizeof(header);
    header.biWidth = width;
    header.biHeight = height;
    header.biPlanes = 1;
    header.biBitCount = 24;
    header.biCompression = BI_RGB;
    const DWORD rowSize = ((width * 3 + 3) & ~3);
    std::vector<unsigned char> pixels(static_cast<size_t>(rowSize) * height);
    BITMAPINFO info{};
    info.bmiHeader = header;
    GetDIBits(memory, bitmap, 0, height, pixels.data(), &info, DIB_RGB_COLORS);
    BITMAPFILEHEADER fileHeader{};
    fileHeader.bfType = 0x4d42;
    fileHeader.bfOffBits = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER);
    fileHeader.bfSize = fileHeader.bfOffBits + static_cast<DWORD>(pixels.size());
    std::ofstream out(fs::path(path), std::ios::binary);
    out.write(reinterpret_cast<const char*>(&fileHeader), sizeof(fileHeader));
    out.write(reinterpret_cast<const char*>(&header), sizeof(header));
    out.write(reinterpret_cast<const char*>(pixels.data()), pixels.size());
    SelectObject(memory, old);
    DeleteObject(bitmap);
    DeleteDC(memory);
    ReleaseDC(window, source);
    return out.good();
}

LRESULT CALLBACK WindowProcedure(HWND window, UINT message, WPARAM wParam, LPARAM lParam) {
    auto* state = reinterpret_cast<AppState*>(GetWindowLongPtrW(window, GWLP_USERDATA));
    switch (message) {
    case WM_CREATE: {
        state = new AppState{};
        state->window = window;
        SetWindowLongPtrW(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(state));
        state->font = CreateFontW(-18, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                  OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                  DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei UI");
        state->titleFont = CreateFontW(-27, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                                       OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                       DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei UI");
        state->statusBrush = CreateSolidBrush(RGB(238, 245, 255));
        state->appIcon = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_ICON,
                                         38, 26, 58, 58, window, reinterpret_cast<HMENU>(IDC_APP_ICON),
                                         GetModuleHandleW(nullptr), nullptr);
        HICON headerIcon = reinterpret_cast<HICON>(LoadImageW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(101),
                                                               IMAGE_ICON, 52, 52, LR_DEFAULTCOLOR));
        SendMessageW(state->appIcon, STM_SETICON, reinterpret_cast<WPARAM>(headerIcon), 0);
        AddStatic(window, L"BHH PDF 转 Word", 128, 34, 560, 40, state->titleFont, IDC_TITLE);
        AddStatic(window, L"智能排版 · 数字校验 · 表格可编辑 · 单文件离线", 130, 76, 590, 24, state->font, IDC_SUBTITLE);
        AddStatic(window, L"选择 PDF 文件", 48, 146, 180, 22, state->font);
        AddStatic(window, L"Word 保存位置", 48, 236, 180, 22, state->font);
        AddStatic(window, L"转换方式", 48, 326, 150, 22, state->font);
        AddStatic(window, L"运行状态", 48, 498, 150, 24, state->font, IDC_STATUS_LABEL);
        state->input = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                       0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_INPUT), GetModuleHandleW(nullptr), nullptr);
        CreateWindowExW(0, L"BUTTON", L"选择 PDF", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                        0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_INPUT_BROWSE), GetModuleHandleW(nullptr), nullptr);
        state->output = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                        0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_OUTPUT), GetModuleHandleW(nullptr), nullptr);
        CreateWindowExW(0, L"BUTTON", L"保存位置", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                        0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_OUTPUT_BROWSE), GetModuleHandleW(nullptr), nullptr);
        state->engine = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
                                        0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_ENGINE), GetModuleHandleW(nullptr), nullptr);
        SendMessageW(state->engine, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"自动选择（推荐：内置智能可编辑引擎）"));
        SendMessageW(state->engine, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"智能可编辑（版式重建、表格识别、数字校验）"));
        SendMessageW(state->engine, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"精准版式（整页图片，不可编辑）"));
        SendMessageW(state->engine, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Word 高保真可编辑（需要 Microsoft Word）"));
        SendMessageW(state->engine, CB_SETCURSEL, 0, 0);
        state->convert = CreateWindowExW(0, L"BUTTON", L"开始转换", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                         0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_CONVERT), GetModuleHandleW(nullptr), nullptr);
        state->openFolder = CreateWindowExW(0, L"BUTTON", L"打开目录", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                            0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_OPEN_FOLDER), GetModuleHandleW(nullptr), nullptr);
        state->status = CreateWindowExW(0, L"STATIC",
                                        WordAvailable() ? L"就绪  ·  自动模式使用内置智能引擎；Word 仅供手动高保真模式使用。\r\n提示  ·  选择 PDF 或直接拖入窗口即可开始。"
                                                        : L"就绪  ·  自动模式使用内置智能引擎重建版式、表格并校验数字。\r\n提示  ·  选择 PDF 或直接拖入窗口即可开始。",
                                        WS_CHILD | WS_VISIBLE | SS_LEFT,
                                        0, 0, 0, 0, window, reinterpret_cast<HMENU>(IDC_STATUS), GetModuleHandleW(nullptr), nullptr);
        for (int id = IDC_INPUT; id <= IDC_STATUS; ++id) {
            HWND control = GetDlgItem(window, id);
            if (control) SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(state->font), TRUE);
        }
        DragAcceptFiles(window, TRUE);
        RECT client{};
        GetClientRect(window, &client);
        LayoutControls(state, client.right, client.bottom);
        if (!g_screenshotPath.empty()) SetTimer(window, 1, 700, nullptr);
        return 0;
    }
    case WM_SIZE:
        if (state) LayoutControls(state, LOWORD(lParam), HIWORD(lParam));
        InvalidateRect(window, nullptr, TRUE);
        return 0;
    case WM_GETMINMAXINFO: {
        auto* info = reinterpret_cast<MINMAXINFO*>(lParam);
        info->ptMinTrackSize.x = 800;
        info->ptMinTrackSize.y = 680;
        return 0;
    }
    case WM_ERASEBKGND:
        DrawBackground(window, reinterpret_cast<HDC>(wParam));
        return 1;
    case WM_PAINT: {
        PAINTSTRUCT paint{};
        HDC dc = BeginPaint(window, &paint);
        DrawBackground(window, dc);
        EndPaint(window, &paint);
        return 0;
    }
    case WM_CTLCOLORSTATIC: {
        HDC dc = reinterpret_cast<HDC>(wParam);
        SetBkMode(dc, TRANSPARENT);
        SetTextColor(dc, RGB(17, 24, 39));
        HWND control = reinterpret_cast<HWND>(lParam);
        if (state && GetDlgCtrlID(control) == IDC_STATUS) {
            SetBkMode(dc, OPAQUE);
            SetBkColor(dc, RGB(239, 246, 255));
            SetTextColor(dc, RGB(22, 70, 166));
            return reinterpret_cast<LRESULT>(state->statusBrush);
        }
        const int id = GetDlgCtrlID(control);
        if (id == IDC_TITLE) {
            SetTextColor(dc, RGB(255, 255, 255));
        } else if (id == IDC_SUBTITLE) {
            SetTextColor(dc, RGB(199, 224, 255));
        }
        return reinterpret_cast<LRESULT>(GetStockObject(NULL_BRUSH));
    }
    case WM_DRAWITEM:
        DrawOwnerButton(reinterpret_cast<const DRAWITEMSTRUCT*>(lParam));
        return TRUE;
    case WM_DROPFILES:
        if (state && !state->busy) {
            wchar_t path[32768]{};
            if (DragQueryFileW(reinterpret_cast<HDROP>(wParam), 0, path, static_cast<UINT>(std::size(path)))) {
                fs::path source(path);
                std::wstring ext = source.extension().wstring();
                std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
                if (ext == L".pdf") {
                    SetWindowTextW(state->input, path);
                    source.replace_extension(L".docx");
                    SetWindowTextW(state->output, source.c_str());
                } else {
                    MessageBoxW(window, L"请拖入 PDF 文件。", L"文件类型不支持", MB_OK | MB_ICONWARNING);
                }
            }
            DragFinish(reinterpret_cast<HDROP>(wParam));
        }
        return 0;
    case WM_COMMAND:
        if (!state) break;
        switch (LOWORD(wParam)) {
        case IDC_INPUT_BROWSE: ChooseInput(state); return 0;
        case IDC_OUTPUT_BROWSE: ChooseOutput(state); return 0;
        case IDC_CONVERT: StartConversion(state); return 0;
        case IDC_OPEN_FOLDER: {
            fs::path output(GetText(state->output));
            fs::path folder = output.has_parent_path() ? output.parent_path() : fs::current_path();
            ShellExecuteW(window, L"open", folder.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
            return 0;
        }
        }
        break;
    case WM_STATUS_TEXT:
        if (state) {
            std::unique_ptr<std::wstring> text(reinterpret_cast<std::wstring*>(lParam));
            SetWindowTextW(state->status, text->c_str());
        }
        return 0;
    case WM_CONVERSION_DONE:
        if (state) {
            std::unique_ptr<std::wstring> text(reinterpret_cast<std::wstring*>(lParam));
            state->busy = false;
            EnableWindow(state->convert, TRUE);
            EnableWindow(state->input, TRUE);
            EnableWindow(state->output, TRUE);
            SetWindowTextW(state->status, text->c_str());
            MessageBoxW(window, text->c_str(), wParam ? L"转换完成" : L"转换失败",
                        MB_OK | (wParam ? MB_ICONINFORMATION : MB_ICONERROR));
        }
        return 0;
    case WM_TIMER:
        if (wParam == 1 && !g_screenshotPath.empty()) {
            KillTimer(window, 1);
            CaptureWindowBmp(window, g_screenshotPath);
            PostMessageW(window, WM_CLOSE, 0, 0);
        }
        return 0;
    case WM_CLOSE:
        if (state && state->busy) {
            MessageBoxW(window, L"转换仍在进行，请等待完成后再关闭。", L"正在转换", MB_OK | MB_ICONINFORMATION);
            return 0;
        }
        DestroyWindow(window);
        return 0;
    case WM_DESTROY:
        CleanupEmbeddedTools();
        if (state) {
            DeleteObject(state->font);
            DeleteObject(state->titleFont);
            DeleteObject(state->statusBrush);
            delete state;
        }
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(window, message, wParam, lParam);
}

void ParseArguments() {
    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (!argv) return;
    for (int i = 1; i < argc; ++i) {
        const std::wstring arg = argv[i];
        if (arg == L"--convert" && i + 2 < argc) {
            g_cliMode = true;
            g_cliInput = argv[++i];
            g_cliOutput = argv[++i];
        } else if (arg == L"--engine" && i + 1 < argc) {
            const std::wstring value = argv[++i];
            g_cliEngine = value == L"word" ? 3 : (value == L"image" ? 2 : (value == L"auto" ? 0 : 1));
        } else if (arg == L"--screenshot" && i + 1 < argc) {
            g_screenshotPath = argv[++i];
        }
    }
    LocalFree(argv);
}

} // namespace

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE, PWSTR, int showCommand) {
    ParseArguments();
    if (g_cliMode) {
        ConversionJob job{nullptr, g_cliInput, g_cliOutput, g_cliEngine};
        std::wstring message;
        const bool ok = PerformConversion(job, message);
        CleanupEmbeddedTools();
        return ok ? 0 : 1;
    }
    SetProcessDPIAware();
    INITCOMMONCONTROLSEX controls{sizeof(controls), ICC_STANDARD_CLASSES};
    InitCommonControlsEx(&controls);
    WNDCLASSEXW windowClass{};
    windowClass.cbSize = sizeof(windowClass);
    windowClass.lpfnWndProc = WindowProcedure;
    windowClass.hInstance = instance;
    windowClass.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    windowClass.hIcon = LoadIconW(instance, MAKEINTRESOURCEW(101));
    windowClass.hIconSm = reinterpret_cast<HICON>(LoadImageW(instance, MAKEINTRESOURCEW(101), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR));
    windowClass.hbrBackground = nullptr;
    windowClass.lpszClassName = kClassName;
    if (!RegisterClassExW(&windowClass)) return 1;
    HWND window = CreateWindowExW(0, kClassName, L"BHH PDF 转 Word",
                                   WS_OVERLAPPEDWINDOW & ~WS_MAXIMIZEBOX,
                                   CW_USEDEFAULT, CW_USEDEFAULT, 860, 700,
                                   nullptr, nullptr, instance, nullptr);
    if (!window) return 1;
    ShowWindow(window, showCommand);
    UpdateWindow(window);
    MSG message{};
    while (GetMessageW(&message, nullptr, 0, 0) > 0) {
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }
    return static_cast<int>(message.wParam);
}
