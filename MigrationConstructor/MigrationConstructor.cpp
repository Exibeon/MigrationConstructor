#include <windows.h>
#include <commctrl.h>
#include <string>
#include <vector>
#include <fstream>
#include <sstream>
#include <iomanip>
#include <algorithm>
#include <cwctype>
#include <memory>

#pragma comment(lib, "comctl32.lib")
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

// Константы и определения
namespace {
    constexpr int ID_COMBO_BASE = 101;
    constexpr int ID_TEXTBOX = 117;
    constexpr int ID_BUTTON = 118;
    constexpr int ID_ID_EDIT = 119;
    constexpr int ID_LOGIN_EDIT = 120;
    constexpr int ID_CLEAR_BUTTON = 121;
    constexpr int ID_ADD_FIELD_BUTTON = 122;
    constexpr int ID_EXTRA_FIELD_BASE = 123;
    constexpr int MAX_EXTRA_FIELDS = 3;
    constexpr int COMBO_COLUMNS = 4;
    constexpr int DEFAULT_MARGIN = 5;
    constexpr int COMBO_HEIGHT = 50;
    constexpr int LABEL_HEIGHT = 20;
    constexpr int EDIT_HEIGHT = 22;
    constexpr int BUTTON_HEIGHT = 30;
    constexpr int BUTTON_WIDTH = 150;
    constexpr int TEXTBOX_HEIGHT = 200;

    const std::vector<std::wstring> comboBoxFiles = {
        L"role.txt", L"headLead.txt", L"fullname.txt", L"position.txt",
        L"department.txt", L"protectedInfoAccess.txt", L"desks.txt", L"region.txt",
        L"personalNumber.txt", L"pointOfSale.txt", L"coordinator.txt",
        L"BaseMarketFinanceSectors.txt", L"CredDocInvestOperatons.txt",
        L"middleOfficeSub.txt", L"lawyersSub.txt", L"percIndCoordinator.txt"
    };
}

// Структура для хранения состояния приложения
struct AppState {
    int idCounter = 1;
    int loginCounter = 1;
    std::vector<HWND> comboBoxes;
    std::vector<HWND> comboLabels;
    std::vector<HWND> extraFields;
    std::vector<HWND> extraLabels;
    HWND hText = nullptr;
    HWND hIdEdit = nullptr;
    HWND hLoginEdit = nullptr;
    int extraFieldsCount = 0;
    HFONT hFont = nullptr;

    ~AppState() {
        if (hFont) DeleteObject(hFont);
        for (HWND hCombo : comboBoxes) {
            auto* pItems = reinterpret_cast<std::vector<std::wstring>*>(GetWindowLongPtr(hCombo, GWLP_USERDATA));
            delete pItems;
        }
    }
};

// Вспомогательные функции
namespace {
    bool StartsWithCaseInsensitive(const std::wstring& str, const std::wstring& prefix) {
        if (prefix.empty()) return true;
        if (str.length() < prefix.length()) return false;

        return std::equal(prefix.begin(), prefix.end(), str.begin(), [](wchar_t ch1, wchar_t ch2) {
            return towupper(ch1) == towupper(ch2);
            });
    }

    std::wstring GetWindowTextStr(HWND hWnd) {
        const int length = GetWindowTextLength(hWnd) + 1;
        std::unique_ptr<wchar_t[]> buffer(new wchar_t[length]);
        GetWindowText(hWnd, buffer.get(), length);
        return std::wstring(buffer.get());
    }

    void SetWindowTextStr(HWND hWnd, const std::wstring& text) {
        SetWindowText(hWnd, text.c_str());
    }

    std::wstring GetComboBoxLabel(const std::wstring& filename) {
        size_t lastdot = filename.find_last_of(L".");
        return (lastdot == std::wstring::npos) ? filename : filename.substr(0, lastdot);
    }

    std::vector<std::wstring> ReadFileToVector(const std::wstring& filename) {
        std::vector<std::wstring> items;
        std::ifstream file(filename, std::ios::binary);

        if (!file.is_open()) {
            MessageBoxW(NULL, (L"Ошибка открытия файла: " + filename).c_str(), L"Ошибка", MB_ICONERROR);
            return items;
        }

        // Проверка BOM для UTF-8
        char bom[3] = { 0 };
        file.read(bom, 3);
        if (bom[0] != '\xEF' || bom[1] != '\xBB' || bom[2] != '\xBF') {
            file.seekg(0);
        }

        std::string line;
        while (std::getline(file, line)) {
            if (!line.empty()) {
                line.erase(std::remove(line.begin(), line.end(), '\r'), line.end());

                int wcharsCount = MultiByteToWideChar(CP_UTF8, 0, line.c_str(), -1, NULL, 0);
                if (wcharsCount > 0) {
                    std::wstring wline(wcharsCount, 0);
                    MultiByteToWideChar(CP_UTF8, 0, line.c_str(), -1, &wline[0], wcharsCount);
                    wline.pop_back();
                    items.push_back(wline);
                }
            }
        }

        return items;
    }

    void LoadComboBox(HWND hCombo, const std::wstring& filename) {
        auto items = ReadFileToVector(filename);
        auto* pOriginalItems = new std::vector<std::wstring>(std::move(items));
        SetWindowLongPtr(hCombo, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pOriginalItems));

        SendMessage(hCombo, CB_RESETCONTENT, 0, 0);
        for (const auto& item : *pOriginalItems) {
            SendMessage(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item.c_str()));
        }
    }

    void FilterComboBox(HWND hCombo, const std::wstring& filter) {
        auto* pOriginalItems = reinterpret_cast<std::vector<std::wstring>*>(GetWindowLongPtr(hCombo, GWLP_USERDATA));
        if (!pOriginalItems) return;

        DWORD startPos, endPos;
        SendMessage(hCombo, CB_GETEDITSEL, reinterpret_cast<WPARAM>(&startPos), reinterpret_cast<LPARAM>(&endPos));

        SendMessage(hCombo, CB_RESETCONTENT, 0, 0);

        bool hasMatches = false;
        for (const auto& item : *pOriginalItems) {
            if (StartsWithCaseInsensitive(item, filter)) {
                SendMessage(hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(item.c_str()));
                hasMatches = true;
            }
        }

        SetWindowText(hCombo, filter.c_str());
        SendMessage(hCombo, CB_SETEDITSEL, 0, MAKELPARAM(startPos, endPos));

        if (hasMatches) {
            SendMessage(hCombo, CB_SHOWDROPDOWN, TRUE, 0);
        }
    }

    void AppendTextToEdit(HWND hEdit, const std::wstring& text) {
        const int len = GetWindowTextLength(hEdit);
        SendMessage(hEdit, EM_SETSEL, len, len);
        if (len > 0) SendMessage(hEdit, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(L"\r\n"));
        SendMessage(hEdit, EM_REPLACESEL, FALSE, reinterpret_cast<LPARAM>(text.c_str()));
        SendMessage(hEdit, EM_SCROLLCARET, 0, 0);
    }

    HWND CreateLabel(HWND hParent, const std::wstring& text, int x, int y, int width, HFONT hFont) {
        HWND hLabel = CreateWindow(WC_STATIC, text.c_str(),
            WS_VISIBLE | WS_CHILD,
            x, y, width, LABEL_HEIGHT, hParent, NULL, NULL, NULL);
        SendMessage(hLabel, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
        return hLabel;
    }

    HWND CreateComboBox(HWND hParent, int x, int y, int width, int id, HFONT hFont) {
        HWND hCombo = CreateWindow(WC_COMBOBOX, L"",
            CBS_DROPDOWN | CBS_HASSTRINGS | WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | CBS_AUTOHSCROLL,
            x, y, width, COMBO_HEIGHT, hParent, reinterpret_cast<HMENU>(id), NULL, NULL);
        SendMessage(hCombo, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
        return hCombo;
    }

    HWND CreateEdit(HWND hParent, int x, int y, int width, int id, HFONT hFont, DWORD extraStyles = 0) {
        HWND hEdit = CreateWindow(WC_EDIT, L"",
            WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL | extraStyles,
            x, y, width, EDIT_HEIGHT, hParent, reinterpret_cast<HMENU>(id), NULL, NULL);
        SendMessage(hEdit, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
        return hEdit;
    }

    HWND CreateButton(HWND hParent, const std::wstring& text, int x, int y, int id, HFONT hFont) {
        HWND hButton = CreateWindow(WC_BUTTON, text.c_str(),
            WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            x, y, BUTTON_WIDTH, BUTTON_HEIGHT, hParent, reinterpret_cast<HMENU>(id), NULL, NULL);
        SendMessage(hButton, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);
        return hButton;
    }
}

// Основные функции приложения
namespace {
    void UpdateTextBox(AppState* state) {
        if (!state || !state->hLoginEdit || !state->hText || !state->hIdEdit) return;

        const int currentId = std::max<int>(1, _wtoi(GetWindowTextStr(state->hIdEdit).c_str()));
        std::wstring loginBase = GetWindowTextStr(state->hLoginEdit);

        // Обработка суффикса логина
        const size_t underscorePos = loginBase.find_last_of(L'_');
        if (underscorePos != std::wstring::npos &&
            loginBase.length() - underscorePos == 3 &&
            iswdigit(loginBase[underscorePos + 1]) &&
            iswdigit(loginBase[underscorePos + 2])) {
            loginBase = loginBase.substr(0, underscorePos);
        }

        // Формирование логина
        std::wstringstream loginSuffix;
        loginSuffix << L"_" << std::setw(2) << std::setfill(L'0') << state->loginCounter;
        const std::wstring fullLogin = loginBase + loginSuffix.str();

        // Формирование строки результата
        std::wstring result = std::to_wstring(currentId) + L";" + fullLogin + L";";

        // Добавление значений из комбобоксов
        for (size_t i = 0; i < state->comboBoxes.size(); ++i) {
            if (state->comboBoxes[i]) {
                result += GetWindowTextStr(state->comboBoxes[i]);
                if (i < state->comboBoxes.size() - 1 || !state->extraFields.empty()) {
                    result += L";";
                }
            }
        }

        // Добавление значений из дополнительных полей
        for (size_t i = 0; i < state->extraFields.size(); ++i) {
            if (state->extraFields[i]) {
                result += GetWindowTextStr(state->extraFields[i]);
                if (i < state->extraFields.size() - 1) {
                    result += L";";
                }
            }
        }

        AppendTextToEdit(state->hText, result);

        // Обновление счетчиков
        state->idCounter = currentId + 1;
        state->loginCounter++;

        SetWindowTextStr(state->hIdEdit, std::to_wstring(state->idCounter));
        SetWindowTextStr(state->hLoginEdit, loginBase);
    }

    void ResetCounters(AppState* state) {
        if (!state) return;

        state->idCounter = 1;
        state->loginCounter = 1;
        SetWindowTextStr(state->hText, L"");
        SetWindowTextStr(state->hIdEdit, L"1");
        SetWindowTextStr(state->hLoginEdit, L"user");
    }

    void AddExtraField(AppState* state, HWND hWnd) {
        if (!state || state->extraFieldsCount >= MAX_EXTRA_FIELDS) return;

        const int totalControls = state->comboBoxes.size() + state->extraFieldsCount;
        const int col = totalControls % COMBO_COLUMNS;
        const int row = totalControls / COMBO_COLUMNS;

        const int x = DEFAULT_MARGIN + col * (BUTTON_WIDTH + DEFAULT_MARGIN);
        const int y = 40 + row * (COMBO_HEIGHT + LABEL_HEIGHT + 2);

        // Создание метки
        const std::wstring label = L"Доп. поле " + std::to_wstring(state->extraFieldsCount + 1);
        HWND hLabel = CreateLabel(hWnd, label, x, y, BUTTON_WIDTH, state->hFont);
        state->extraLabels.push_back(hLabel);

        // Создание поля ввода
        HWND hField = CreateEdit(hWnd, x, y + LABEL_HEIGHT + 2, BUTTON_WIDTH,
            ID_EXTRA_FIELD_BASE + state->extraFieldsCount, state->hFont);
        state->extraFields.push_back(hField);
        state->extraFieldsCount++;

        // Обновление размера окна
        RECT rc;
        GetClientRect(hWnd, &rc);
        SendMessage(hWnd, WM_SIZE, SIZE_RESTORED, MAKELPARAM(rc.right - rc.left, rc.bottom - rc.top));
    }

    void PositionControls(AppState* state, HWND hWnd, int width) {
        if (!state) return;

        const int comboWidth = (width - DEFAULT_MARGIN * (COMBO_COLUMNS + 1)) / COMBO_COLUMNS;
        const int effectiveComboWidth = std::max<int>(comboWidth, 150);

        // Позиционирование комбобоксов и дополнительных полей
        for (size_t i = 0; i < state->comboBoxes.size() + state->extraFields.size(); ++i) {
            const int col = i % COMBO_COLUMNS;
            const int row = i / COMBO_COLUMNS;
            const int x = DEFAULT_MARGIN + col * (effectiveComboWidth + DEFAULT_MARGIN);
            int y = 40 + row * (COMBO_HEIGHT + LABEL_HEIGHT + 15);

            if (i < state->comboBoxes.size()) {
                if (i < state->comboLabels.size()) {
                    SetWindowPos(state->comboLabels[i], NULL, x, y, effectiveComboWidth, LABEL_HEIGHT, SWP_NOZORDER);
                }
                y += LABEL_HEIGHT + 2;
                SetWindowPos(state->comboBoxes[i], NULL, x, y, effectiveComboWidth, COMBO_HEIGHT, SWP_NOZORDER);
            }
            else {
                const size_t extraIdx = i - state->comboBoxes.size();
                if (extraIdx < state->extraLabels.size()) {
                    SetWindowPos(state->extraLabels[extraIdx], NULL, x, y, effectiveComboWidth, LABEL_HEIGHT, SWP_NOZORDER);
                }
                y += LABEL_HEIGHT + 2;
                if (extraIdx < state->extraFields.size()) {
                    SetWindowPos(state->extraFields[extraIdx], NULL, x, y, effectiveComboWidth, EDIT_HEIGHT, SWP_NOZORDER);
                }
            }
        }

        // Позиционирование текстового поля
        const int lastRow = (state->comboBoxes.size() + state->extraFields.size() + COMBO_COLUMNS - 1) / COMBO_COLUMNS;
        const int textBoxTop = 40 + lastRow * (COMBO_HEIGHT + LABEL_HEIGHT + 15) + 20;

        if (state->hText) {
            SetWindowPos(state->hText, NULL, DEFAULT_MARGIN, textBoxTop,
                width - 2 * DEFAULT_MARGIN, TEXTBOX_HEIGHT, SWP_NOZORDER);
        }

        // Позиционирование кнопок
        const int buttonsY = 750 - 50;
        SetWindowPos(GetDlgItem(hWnd, ID_BUTTON), NULL,
            DEFAULT_MARGIN, buttonsY, BUTTON_WIDTH, BUTTON_HEIGHT, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hWnd, ID_CLEAR_BUTTON), NULL,
            DEFAULT_MARGIN + 160, buttonsY, BUTTON_WIDTH, BUTTON_HEIGHT, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hWnd, ID_ADD_FIELD_BUTTON), NULL,
            DEFAULT_MARGIN + 320, buttonsY, BUTTON_WIDTH, BUTTON_HEIGHT, SWP_NOZORDER);
    }
}

// Оконная процедура
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    auto* pState = reinterpret_cast<AppState*>(GetWindowLongPtr(hWnd, GWLP_USERDATA));

    switch (message) {
    case WM_CREATE: {
        pState = new AppState();
        SetWindowLongPtr(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pState));

        INITCOMMONCONTROLSEX icex = { sizeof(INITCOMMONCONTROLSEX), ICC_STANDARD_CLASSES | ICC_LISTVIEW_CLASSES };
        InitCommonControlsEx(&icex);

        // Создание шрифта
        pState->hFont = CreateFont(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");

        // Создание элементов управления ID и Login
        CreateLabel(hWnd, L"ID:", DEFAULT_MARGIN, 10, 30, pState->hFont);
        pState->hIdEdit = CreateEdit(hWnd, DEFAULT_MARGIN + 35, 10, 120, ID_ID_EDIT, pState->hFont, ES_NUMBER);
        SetWindowTextStr(pState->hIdEdit, L"1");

        CreateLabel(hWnd, L"Login:", DEFAULT_MARGIN + 200, 10, 50, pState->hFont);
        pState->hLoginEdit = CreateEdit(hWnd, DEFAULT_MARGIN + 255, 10, 120, ID_LOGIN_EDIT, pState->hFont);
        SetWindowTextStr(pState->hLoginEdit, L"user");

        // Создание комбобоксов
        for (size_t i = 0; i < comboBoxFiles.size(); ++i) {
            const int col = i % COMBO_COLUMNS;
            const int row = i / COMBO_COLUMNS;
            const int x = DEFAULT_MARGIN + col * (BUTTON_WIDTH + DEFAULT_MARGIN);
            const int y = 40 + row * (COMBO_HEIGHT + LABEL_HEIGHT + 15);

            // Метка
            std::wstring label = GetComboBoxLabel(comboBoxFiles[i]);
            HWND hLabel = CreateLabel(hWnd, label, x, y, BUTTON_WIDTH, pState->hFont);
            pState->comboLabels.push_back(hLabel);

            // Комбобокс
            HWND hCombo = CreateComboBox(hWnd, x, y + LABEL_HEIGHT + 2, BUTTON_WIDTH,
                ID_COMBO_BASE + i, pState->hFont);
            pState->comboBoxes.push_back(hCombo);
            LoadComboBox(hCombo, comboBoxFiles[i]);
        }

        // Текстовое поле
        const int lastRow = (comboBoxFiles.size() + COMBO_COLUMNS - 1) / COMBO_COLUMNS;
        const int textBoxTop = 40 + lastRow * (COMBO_HEIGHT + LABEL_HEIGHT + 15) + 20;

        pState->hText = CreateWindow(WC_EDIT, L"",
            WS_VISIBLE | WS_CHILD | WS_BORDER | ES_MULTILINE | ES_AUTOVSCROLL |
            ES_AUTOHSCROLL | WS_VSCROLL | WS_HSCROLL | ES_WANTRETURN,
            DEFAULT_MARGIN, textBoxTop, BUTTON_WIDTH * COMBO_COLUMNS + DEFAULT_MARGIN * (COMBO_COLUMNS - 1),
            TEXTBOX_HEIGHT, hWnd, reinterpret_cast<HMENU>(ID_TEXTBOX), NULL, NULL);
        SendMessage(pState->hText, WM_SETFONT, reinterpret_cast<WPARAM>(pState->hFont), TRUE);

        // Кнопки
        CreateButton(hWnd, L"Добавить запись", DEFAULT_MARGIN, 700, ID_BUTTON, pState->hFont);
        CreateButton(hWnd, L"Очистка", DEFAULT_MARGIN + 160, 700, ID_CLEAR_BUTTON, pState->hFont);
        CreateButton(hWnd, L"Добавить поле", DEFAULT_MARGIN + 320, 700, ID_ADD_FIELD_BUTTON, pState->hFont);

        SetWindowPos(hWnd, NULL, 0, 0, 800, 800, SWP_NOMOVE | SWP_NOZORDER);
        break;
    }

    case WM_SIZE:
        if (pState) {
            PositionControls(pState, hWnd, LOWORD(lParam));
            InvalidateRect(hWnd, NULL, TRUE);
            UpdateWindow(hWnd);
        }
        break;

    case WM_COMMAND:
        if (HIWORD(wParam) == CBN_EDITUPDATE) {
            HWND hCombo = reinterpret_cast<HWND>(lParam);
            if (hCombo && (GetWindowLongPtr(hCombo, GWL_STYLE) & CBS_DROPDOWN)) {
                FilterComboBox(hCombo, GetWindowTextStr(hCombo));
            }
        }
        else {
            switch (LOWORD(wParam)) {
            case ID_BUTTON: UpdateTextBox(pState); break;
            case ID_CLEAR_BUTTON: ResetCounters(pState); break;
            case ID_ADD_FIELD_BUTTON: AddExtraField(pState, hWnd); break;
            }
        }
        break;

    case WM_DESTROY:
        delete pState;
        SetWindowLongPtr(hWnd, GWLP_USERDATA, 0);
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow) {
    WNDCLASS wc = { 0 };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    wc.lpszClassName = L"DropdownApp";
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);

    if (!RegisterClass(&wc)) {
        MessageBox(NULL, L"Ошибка регистрации класса окна!", L"Ошибка", MB_ICONERROR);
        return 0;
    }

    HWND hWnd = CreateWindow(wc.lpszClassName, L"MigrationConstructor",
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        CW_USEDEFAULT, CW_USEDEFAULT, 800, 750,
        NULL, NULL, hInstance, NULL);

    if (!hWnd) {
        MessageBox(NULL, L"Ошибка создания окна!", L"Ошибка", MB_ICONERROR);
        return 0;
    }

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return static_cast<int>(msg.wParam);
}
