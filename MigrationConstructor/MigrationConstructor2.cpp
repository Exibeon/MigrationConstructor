#include <windows.h>
#include <commctrl.h>
#include <string>
#include <vector>
#include <fstream>
#include <sstream>
#include <iomanip>
#include <algorithm>
#include <cwctype>

#pragma comment(lib, "comctl32.lib")
#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#define ID_COMBO_BASE 101
#define ID_TEXTBOX 117
#define ID_BUTTON 118
#define ID_ID_EDIT 119
#define ID_LOGIN_EDIT 120
#define ID_CLEAR_BUTTON 121
#define ID_ADD_FIELD_BUTTON 122
#define ID_EXTRA_FIELD_BASE 123

const std::vector<std::wstring> comboBoxFiles = {
    L"role.txt", L"headLead.txt", L"fullname.txt", L"position.txt",
    L"department.txt", L"protectedInfoAccess.txt", L"desks.txt", L"region.txt",
    L"personalNumber.txt", L"pointOfSale.txt", L"coordinator.txt",
    L"BaseMarketFinanceSectors.txt", L"CredDocInvestOperatons.txt",
    L"middleOfficeSub.txt", L"lawyersSub.txt", L"percIndCoordinator.txt"
};

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
};

bool StartsWithCaseInsensitive(const std::wstring& str, const std::wstring& prefix) {
    if (prefix.empty()) return true;
    if (str.length() < prefix.length()) return false;

    return std::equal(
        prefix.begin(), prefix.end(),
        str.begin(),
        [](wchar_t ch1, wchar_t ch2) {
            return towupper(ch1) == towupper(ch2);
        }
    );
}

std::wstring GetComboBoxText(HWND hCombo) {
    wchar_t buffer[256] = { 0 };
    GetWindowText(hCombo, buffer, 256);
    return std::wstring(buffer);
}

void FilterComboBox(HWND hCombo, const std::wstring& filter) {
    std::vector<std::wstring>* pOriginalItems = reinterpret_cast<std::vector<std::wstring>*>(
        GetWindowLongPtr(hCombo, GWLP_USERDATA));

    if (!pOriginalItems) return;

    DWORD startPos, endPos;
    SendMessage(hCombo, CB_GETEDITSEL, (WPARAM)&startPos, (LPARAM)&endPos);

    SendMessage(hCombo, CB_RESETCONTENT, 0, 0);

    bool hasMatches = false;
    for (const auto& item : *pOriginalItems) {
        if (StartsWithCaseInsensitive(item, filter)) {
            SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)item.c_str());
            hasMatches = true;
        }
    }

    SetWindowText(hCombo, filter.c_str());
    SendMessage(hCombo, CB_SETEDITSEL, 0, MAKELPARAM(startPos, endPos));

    if (hasMatches) {
        SendMessage(hCombo, CB_SHOWDROPDOWN, TRUE, 0);
    }
}

std::wstring GetComboBoxLabel(const std::wstring& filename) {
    size_t lastdot = filename.find_last_of(L".");
    if (lastdot == std::wstring::npos) return filename;
    return filename.substr(0, lastdot);
}

std::vector<std::wstring> ReadFileToVector(const std::wstring& filename) {
    std::vector<std::wstring> items;
    std::ifstream file(filename, std::ios::binary); // Открываем в бинарном режиме

    if (!file.is_open()) {
        MessageBoxW(NULL, (L"Ошибка открытия файла: " + filename).c_str(), L"Ошибка", MB_ICONERROR);
        return items;
    }

    // Проверяем наличие BOM (EF BB BF для UTF-8)
    char bom[3] = { 0 };
    file.read(bom, 3);
    if (bom[0] != '\xEF' || bom[1] != '\xBB' || bom[2] != '\xBF') {
        file.seekg(0); // Если BOM нет, возвращаемся в начало файла
    }

    std::string line;
    while (std::getline(file, line)) {
        if (!line.empty()) {
            // Удаляем \r (если есть)
            line.erase(std::remove(line.begin(), line.end(), '\r'), line.end());

            // Конвертируем UTF-8 → UTF-16 (wstring)
            int wcharsCount = MultiByteToWideChar(CP_UTF8, 0, line.c_str(), -1, NULL, 0);
            if (wcharsCount > 0) {
                std::wstring wline(wcharsCount, 0);
                MultiByteToWideChar(CP_UTF8, 0, line.c_str(), -1, &wline[0], wcharsCount);
                wline.pop_back(); // Удаляем завершающий нуль-символ
                items.push_back(wline);
            }
        }
    }

    file.close();
    return items;
}

void LoadComboBox(HWND hCombo, const std::wstring& filename) {
    auto items = ReadFileToVector(filename);

    std::vector<std::wstring>* pOriginalItems = new std::vector<std::wstring>(items);
    SetWindowLongPtr(hCombo, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pOriginalItems));

    SendMessage(hCombo, CB_RESETCONTENT, 0, 0);
    for (const auto& item : items) {
        SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)item.c_str());
    }
}

std::wstring GetEditText(HWND hEdit) {
    wchar_t buffer[256];
    GetWindowText(hEdit, buffer, 256);
    return std::wstring(buffer);
}

void SetEditText(HWND hEdit, const std::wstring& text) {
    SetWindowText(hEdit, text.c_str());
}

void AppendText(HWND hEdit, const std::wstring& text) {
    int len = GetWindowTextLength(hEdit);
    SendMessage(hEdit, EM_SETSEL, len, len);
    if (len > 0) SendMessage(hEdit, EM_REPLACESEL, FALSE, (LPARAM)L"\r\n");
    SendMessage(hEdit, EM_REPLACESEL, FALSE, (LPARAM)text.c_str());
    SendMessage(hEdit, EM_SCROLLCARET, 0, 0);
}

void UpdateTextBox(AppState* state) {
    if (!state || !state->hLoginEdit || !state->hText || !state->hIdEdit) return;

    std::wstring idStr = GetEditText(state->hIdEdit);
    int currentId = _wtoi(idStr.c_str());
    if (currentId <= 0) currentId = 1;

    std::wstring loginBase = GetEditText(state->hLoginEdit);
    size_t underscorePos = loginBase.find_last_of(L'_');
    if (underscorePos != std::wstring::npos) {
        if (loginBase.length() - underscorePos == 3 &&
            iswdigit(loginBase[underscorePos + 1]) &&
            iswdigit(loginBase[underscorePos + 2])) {
            loginBase = loginBase.substr(0, underscorePos);
        }
    }

    std::wstringstream loginSuffix;
    loginSuffix << L"_" << std::setw(2) << std::setfill(L'0') << state->loginCounter;
    std::wstring fullLogin = loginBase + loginSuffix.str();

    std::wstring result;
    result += std::to_wstring(currentId) + L";" + fullLogin + L";";

    for (size_t i = 0; i < state->comboBoxes.size(); ++i) {
        if (state->comboBoxes[i]) {
            result += GetComboBoxText(state->comboBoxes[i]);
            if (i < state->comboBoxes.size() - 1 || !state->extraFields.empty()) {
                result += L";";
            }
        }
    }

    for (size_t i = 0; i < state->extraFields.size(); ++i) {
        if (state->extraFields[i]) {
            result += GetEditText(state->extraFields[i]);
            if (i < state->extraFields.size() - 1) {
                result += L";";
            }
        }
    }

    AppendText(state->hText, result);

    state->idCounter = currentId + 1;
    state->loginCounter++;

    SetEditText(state->hIdEdit, std::to_wstring(state->idCounter));
    SetEditText(state->hLoginEdit, loginBase);
}

void ResetCounters(AppState* state) {
    if (!state) return;

    state->idCounter = 1;
    state->loginCounter = 1;
    SetEditText(state->hText, L"");
    SetEditText(state->hIdEdit, L"1");
    SetEditText(state->hLoginEdit, L"user");
}

void AddExtraField(AppState* state, HWND hWnd) {
    if (!state || state->extraFieldsCount >= 3) return;

    HFONT hFont = CreateFont(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");

    const int margin = 5;
    const int fieldWidth = 180;
    const int fieldHeight = 22;
    const int labelHeight = 20;
    const int verticalSpacing = 2;
    const int columns = 4;
    const int comboHeight = 50;

    int col = (state->comboBoxes.size() + state->extraFieldsCount) % columns;
    int row = (state->comboBoxes.size() + state->extraFieldsCount) / columns;

    int x = margin + col * (fieldWidth + margin);
    int y = 40 + row * (comboHeight + labelHeight + verticalSpacing);

    std::wstring label = L"Доп. поле " + std::to_wstring(state->extraFieldsCount + 1);
    HWND hLabel = CreateWindow(WC_STATIC, label.c_str(),
        WS_VISIBLE | WS_CHILD,
        x, y, fieldWidth, labelHeight, hWnd, NULL, NULL, NULL);
    SendMessage(hLabel, WM_SETFONT, (WPARAM)hFont, TRUE);
    state->extraLabels.push_back(hLabel);

    y += labelHeight + verticalSpacing;
    HWND hField = CreateWindow(WC_EDIT, L"",
        WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
        x, y, fieldWidth, fieldHeight, hWnd, (HMENU)(ID_EXTRA_FIELD_BASE + state->extraFieldsCount), NULL, NULL);
    SendMessage(hField, WM_SETFONT, (WPARAM)hFont, TRUE);

    state->extraFields.push_back(hField);
    state->extraFieldsCount++;

    RECT rc;
    GetClientRect(hWnd, &rc);
    SendMessage(hWnd, WM_SIZE, SIZE_RESTORED, MAKELPARAM(rc.right - rc.left, rc.bottom - rc.top));
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    AppState* pState = reinterpret_cast<AppState*>(GetWindowLongPtr(hWnd, GWLP_USERDATA));

    switch (message) {
    case WM_CREATE: {
        pState = new AppState();
        SetWindowLongPtr(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pState));

        INITCOMMONCONTROLSEX icex = { sizeof(INITCOMMONCONTROLSEX), ICC_STANDARD_CLASSES | ICC_LISTVIEW_CLASSES };
        InitCommonControlsEx(&icex);

        const int comboWidth = 180;
        const int editWidth = 120;
        const int comboHeight = 50;
        const int labelHeight = 20;
        const int margin = 5;
        const int topMargin = 10;
        const int verticalSpacing = 2;
        const int columns = 4;

        HFONT hFont = CreateFont(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            DEFAULT_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");

        int x = margin;
        int y = topMargin;

        CreateWindow(WC_STATIC, L"ID:", WS_VISIBLE | WS_CHILD,
            x, y, 30, labelHeight, hWnd, NULL, NULL, NULL);

        pState->hIdEdit = CreateWindow(WC_EDIT, L"1", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL | ES_NUMBER,
            x + 35, y, editWidth, 22, hWnd, (HMENU)ID_ID_EDIT, NULL, NULL);

        x += editWidth + 45;

        CreateWindow(WC_STATIC, L"Login:", WS_VISIBLE | WS_CHILD,
            x, y, 50, labelHeight, hWnd, NULL, NULL, NULL);

        pState->hLoginEdit = CreateWindow(WC_EDIT, L"user", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
            x + 55, y, editWidth, 22, hWnd, (HMENU)ID_LOGIN_EDIT, NULL, NULL);

        SendMessage(pState->hIdEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessage(pState->hLoginEdit, WM_SETFONT, (WPARAM)hFont, TRUE);

        y += 30;

        for (size_t i = 0; i < comboBoxFiles.size(); ++i) {
            int col = i % columns;
            int row = i / columns;

            x = margin + col * (comboWidth + margin);
            int currentY = y + row * (comboHeight + labelHeight + verticalSpacing);

            std::wstring label = GetComboBoxLabel(comboBoxFiles[i]);
            HWND hLabel = CreateWindow(WC_STATIC, label.c_str(),
                WS_VISIBLE | WS_CHILD,
                x, currentY, comboWidth, labelHeight, hWnd, NULL, NULL, NULL);
            SendMessage(hLabel, WM_SETFONT, (WPARAM)hFont, TRUE);
            pState->comboLabels.push_back(hLabel);

            currentY += labelHeight + verticalSpacing;
            HWND hCombo = CreateWindow(WC_COMBOBOX, L"",
                CBS_DROPDOWN | CBS_HASSTRINGS | WS_CHILD | WS_VISIBLE | WS_TABSTOP | WS_VSCROLL | CBS_AUTOHSCROLL,
                x, currentY, comboWidth, comboHeight, hWnd, (HMENU)(ID_COMBO_BASE + i), NULL, NULL);

            pState->comboBoxes.push_back(hCombo);
            LoadComboBox(hCombo, comboBoxFiles[i]);
            SendMessage(hCombo, WM_SETFONT, (WPARAM)hFont, TRUE);
        }

        int lastRow = (comboBoxFiles.size() + columns - 1) / columns;
        int textBoxTop = y + lastRow * (comboHeight + labelHeight + verticalSpacing) + 20;
        int textBoxHeight = 200;

        pState->hText = CreateWindow(WC_EDIT, L"",
            WS_VISIBLE | WS_CHILD | WS_BORDER | ES_MULTILINE |
            ES_AUTOVSCROLL | ES_AUTOHSCROLL | WS_VSCROLL | WS_HSCROLL | ES_WANTRETURN,
            margin, textBoxTop, comboWidth * columns + margin * (columns - 1), textBoxHeight, hWnd, (HMENU)ID_TEXTBOX, NULL, NULL);
        SendMessage(pState->hText, WM_SETFONT, (WPARAM)hFont, TRUE);

        int buttonsY = 750 - 50;
        CreateWindow(WC_BUTTON, L"Добавить запись",
            WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            margin, buttonsY, 150, 30, hWnd, (HMENU)ID_BUTTON, NULL, NULL);
        SendMessage(GetDlgItem(hWnd, ID_BUTTON), WM_SETFONT, (WPARAM)hFont, TRUE);

        CreateWindow(WC_BUTTON, L"Очистка",
            WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            margin + 160, buttonsY, 150, 30, hWnd, (HMENU)ID_CLEAR_BUTTON, NULL, NULL);
        SendMessage(GetDlgItem(hWnd, ID_CLEAR_BUTTON), WM_SETFONT, (WPARAM)hFont, TRUE);

        CreateWindow(WC_BUTTON, L"Добавить поле",
            WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
            margin + 320, buttonsY, 150, 30, hWnd, (HMENU)ID_ADD_FIELD_BUTTON, NULL, NULL);
        SendMessage(GetDlgItem(hWnd, ID_ADD_FIELD_BUTTON), WM_SETFONT, (WPARAM)hFont, TRUE);

        SetWindowPos(hWnd, NULL, 0, 0, 800, 800, SWP_NOMOVE | SWP_NOZORDER);
        break;
    }
    case WM_SIZE: {
        if (!pState) break;

        int width = LOWORD(lParam);
        int height = HIWORD(lParam);

        const int columns = 4;
        const int margin = 5;
        int comboWidth = (width - margin * (columns + 1)) / columns;
        if (comboWidth < 150) comboWidth = 150;

        const int comboHeight = 50;
        const int labelHeight = 20;
        const int verticalSpacing = 15;
        const int topMargin = 10;

        for (size_t i = 0; i < pState->comboBoxes.size() + pState->extraFields.size(); ++i) {
            int col = i % columns;
            int row = i / columns;

            int x = margin + col * (comboWidth + margin);
            int y = topMargin + 30 + row * (comboHeight + labelHeight + verticalSpacing);

            if (i < pState->comboBoxes.size()) {
                if (i < pState->comboLabels.size()) {
                    SetWindowPos(pState->comboLabels[i], NULL,
                        x, y, comboWidth, labelHeight, SWP_NOZORDER);
                }
                y += labelHeight + verticalSpacing;
                SetWindowPos(pState->comboBoxes[i], NULL,
                    x, y, comboWidth, comboHeight, SWP_NOZORDER);
            }
            else {
                size_t extraIdx = i - pState->comboBoxes.size();
                if (extraIdx < pState->extraLabels.size()) {
                    SetWindowPos(pState->extraLabels[extraIdx], NULL,
                        x, y, comboWidth, labelHeight, SWP_NOZORDER);
                }
                y += labelHeight + verticalSpacing;
                if (extraIdx < pState->extraFields.size()) {
                    SetWindowPos(pState->extraFields[extraIdx], NULL,
                        x, y, comboWidth, 22, SWP_NOZORDER);
                }
            }
        }

        int lastRow = (pState->comboBoxes.size() + pState->extraFields.size() + columns - 1) / columns;
        int textBoxTop = topMargin + 30 + lastRow * (comboHeight + labelHeight + verticalSpacing) + 20;
        int textBoxHeight = 200;

        if (pState->hText) {
            SetWindowPos(pState->hText, NULL,
                margin, textBoxTop, width - 2 * margin, textBoxHeight, SWP_NOZORDER);
        }

        int buttonsY = 750 - 50;
        SetWindowPos(GetDlgItem(hWnd, ID_BUTTON), NULL,
            margin, buttonsY, 150, 30, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hWnd, ID_CLEAR_BUTTON), NULL,
            margin + 160, buttonsY, 150, 30, SWP_NOZORDER);
        SetWindowPos(GetDlgItem(hWnd, ID_ADD_FIELD_BUTTON), NULL,
            margin + 320, buttonsY, 150, 30, SWP_NOZORDER);

        InvalidateRect(hWnd, NULL, TRUE);
        UpdateWindow(hWnd);
        break;
    }
    case WM_COMMAND:
        if (HIWORD(wParam) == CBN_EDITUPDATE) {
            HWND hCombo = reinterpret_cast<HWND>(lParam);
            if (hCombo && (GetWindowLongPtr(hCombo, GWL_STYLE) & CBS_DROPDOWN)) {
                wchar_t buffer[256];
                GetWindowText(hCombo, buffer, 256);
                FilterComboBox(hCombo, buffer);
            }
        }
        else if (LOWORD(wParam) == ID_BUTTON && pState) {
            UpdateTextBox(pState);
        }
        else if (LOWORD(wParam) == ID_CLEAR_BUTTON && pState) {
            ResetCounters(pState);
        }
        else if (LOWORD(wParam) == ID_ADD_FIELD_BUTTON && pState) {
            AddExtraField(pState, hWnd);
        }
        break;
    case WM_DESTROY:
        if (pState) {
            for (HWND hCombo : pState->comboBoxes) {
                std::vector<std::wstring>* pItems = reinterpret_cast<std::vector<std::wstring>*>(
                    GetWindowLongPtr(hCombo, GWLP_USERDATA));
                delete pItems;
            }

            delete pState;
            SetWindowLongPtr(hWnd, GWLP_USERDATA, 0);
        }
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    WNDCLASS wc = { 0 };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = L"DropdownApp";
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);

    if (!RegisterClass(&wc)) {
        MessageBox(NULL, L"Ошибка регистрации класса окна!", L"Ошибка", MB_ICONERROR);
        return 0;
    }

    HWND hWnd = CreateWindow(wc.lpszClassName, L"Ввод данных сотрудника",
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

    return (int)msg.wParam;
}
