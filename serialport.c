// serialport.c : 定义应用程序的入口点。
//

#include <stdio.h>
#include "framework.h"
#include "serialport.h"
#include "dbt.h"
#include <Locale.h>

// 对话框过程函数实现
HWND CbSerialPort; //串口下拉框
HWND CbBaudRate;   //波特率下拉框
HWND CbByteSize;   //数据位下拉框
HWND CbStopBits;   //停止位下拉框
HWND CbParity;     //校验位下拉框
HWND BtnPort;      //打开关闭按钮
HWND BtnClear;     //清空按钮
HWND BtnSend;      //发送按钮
HWND EditReceive;  //接收区
HWND EditSend;     //发送区
HWND TcMsg;        //提示信息

const TCHAR* CbBaudRateItems[] = { TEXT("9600"), TEXT("38400"), TEXT("57600"), TEXT("115200") };
const TCHAR* CbByteSizeItems[] = { TEXT("8"), TEXT("9") };
const TCHAR* CbStopBitsItems[] = { TEXT("1"), TEXT("1.5"), TEXT("2") };
const TCHAR* CbParityItems[] = { TEXT("NO"), TEXT("ODD"), TEXT("EVEN"), TEXT("MARK"), TEXT("SPACE") };
const TCHAR* BtnPortStateText[] = { TEXT("打开串口"), TEXT("关闭串口") };

int BtnPortState = 0; //0:串口关闭状态,1:串口打开状态
HANDLE hComPort; //串口操作句柄
TCHAR SerialportName[8]; //串口名字符串
TCHAR SerialBaudRate[8]; //波特率字符串
DWORD BaudRate; //波特率
TCHAR SerialByteSize[8]; //数据位字符串
BYTE ByteSize; //数据位
TCHAR FormattedString[256]; //提示信息字符串
TCHAR SerialReceiveBuf[1024]; //接收缓冲区

SYSTEMTIME st;
/* 查找有效串口 */
void FindValidSerialPort(HWND hwndDlg)
{
    WPARAM com_idx = 0;
    for (int i = 1; i <= 255; ++i) {
        swprintf_s(SerialportName, 8, L"COM%d", i);
        HANDLE hSerial = CreateFile(
            SerialportName,
            GENERIC_READ | GENERIC_WRITE,
            0,
            NULL,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL,
            NULL
        );
        if (hSerial != INVALID_HANDLE_VALUE) {
            SendMessage(CbSerialPort, CB_INSERTSTRING, com_idx, (LPARAM)SerialportName);
            com_idx++;
            CloseHandle(hSerial); // Don't forget to close the handle when done
        }
        else {
            // Handle the error if the port is not valid
        }
    }
}

/* 将宽字符字符串转换为UTF-8编码的字符串 */
int WideCharToUTF8(const wchar_t* src, char* dst, int dstSize) {
    return WideCharToMultiByte(CP_UTF8, 0, src, -1, dst, dstSize, NULL, NULL);
}

/* 发送数据到串口 */
void SendData(LPCVOID lpBuffer, DWORD nNumberOfBytesToWrite) {
    DWORD bytesWritten;
    char *utf8String =  (char *)malloc(nNumberOfBytesToWrite*2);
    if (utf8String != NULL) {
        int result = WideCharToUTF8((const wchar_t*)lpBuffer, utf8String, nNumberOfBytesToWrite*2);
        if (result > 0) {
            // 转换成功，result是转换后的UTF-8字符串的长度
            if (WriteFile(hComPort, utf8String, result, &bytesWritten, NULL)) {
                GetLocalTime(&st);
                swprintf_s(FormattedString, sizeof(FormattedString) / sizeof(FormattedString[0]), L"%04d-%02d-%02d %02d:%02d:%02d>发送成功",
                    st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
            }
            else
            {
                GetLocalTime(&st);
                swprintf_s(FormattedString, sizeof(FormattedString) / sizeof(FormattedString[0]), L"%04d-%02d-%02d %02d:%02d:%02d>发送失败",
                    st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
            }
            SetWindowText(TcMsg, FormattedString);
        }
        else {
            // 转换失败，可以调用GetLastError获取错误代码
            wprintf(L"Error converting string to UTF-8 (%lu)\n", GetLastError());
        }
    }
    free(utf8String);
}

/* 串口接收数据显示到edit控件 */
void ReadDataAndUpdateEdit(void) {
    char buffer[1024];
    DWORD bytesRead = 0;
    int wideLength = 0;
    wchar_t* wideString = NULL;

    if (BtnPortState == 1)
    {
        if (ReadFile(hComPort, buffer, sizeof(buffer) - 1, &bytesRead, NULL)) {
            if (bytesRead)
            {
                wideLength = MultiByteToWideChar(CP_UTF8, 0, buffer, bytesRead, NULL, 0);
                // 分配缓冲区以存储转换后的宽字符字符串
                wideString = (wchar_t*)malloc((wideLength+1) * sizeof(TCHAR));
                if (wideString)
                {
                    // 执行转换
                    MultiByteToWideChar(CP_UTF8, 0, buffer, bytesRead, wideString, wideLength * 3);
                    wideString[wideLength] = '\0'; // 确保字符串以空字符结尾
                    wprintf(L"%ls\r\n", wideString);
                    // 将光标设置到文本末尾
                    SendMessage(EditReceive, EM_SETSEL, -1, -1);
                    // 追加文本
                    SendMessage(EditReceive, EM_REPLACESEL, 0, (LPARAM)wideString);
                    free(wideString);
                }

            }
        }
    }

}
INT_PTR CALLBACK DialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    INT_PTR ret = (INT_PTR)FALSE;
    switch (uMsg)
    {
    case WM_INITDIALOG:
    {
        // 初始化
        SetTimer(hwndDlg, 1, 1, NULL); //hwndDlg创建一个1ms的定时器
        // 端口号
        CbSerialPort = GetDlgItem(hwndDlg, IDC_SERIALPORT);
        FindValidSerialPort(CbSerialPort);

        //波特率
        CbBaudRate = GetDlgItem(hwndDlg, IDC_BAUD_RATE);
        for (int i = 0; i < sizeof(CbBaudRateItems) / sizeof(CbBaudRateItems[0]); ++i) {
            SendMessage(CbBaudRate, CB_INSERTSTRING, (WPARAM)i, (LPARAM)CbBaudRateItems[i]);
        }
        SendMessage(CbBaudRate, CB_SETCURSEL, (WPARAM)3, (LPARAM)0); // 选择默认项
        //数据位
        CbByteSize = GetDlgItem(hwndDlg, IDC_BYTE_SIZE);
        for (int i = 0; i < sizeof(CbByteSizeItems) / sizeof(CbByteSizeItems[0]); ++i) {
            SendMessage(CbByteSize, CB_INSERTSTRING, (WPARAM)i, (LPARAM)CbByteSizeItems[i]);
        }
        SendMessage(CbByteSize, CB_SETCURSEL, (WPARAM)0, (LPARAM)0); // 选择默认项
        //停止位
        CbStopBits = GetDlgItem(hwndDlg, IDC_STOP_BITS);
        for (int i = 0; i < sizeof(CbStopBitsItems) / sizeof(CbStopBitsItems[0]); ++i) {
            SendMessage(CbStopBits, CB_INSERTSTRING, (WPARAM)i, (LPARAM)CbStopBitsItems[i]);
        }
        SendMessage(CbStopBits, CB_SETCURSEL, (WPARAM)0, (LPARAM)0); // 选择默认项
        //校验位
        CbParity = GetDlgItem(hwndDlg, IDC_PARITY);
        for (int i = 0; i < sizeof(CbParityItems) / sizeof(CbParityItems[0]); ++i) {
            SendMessage(CbParity, CB_INSERTSTRING, (WPARAM)i, (LPARAM)CbParityItems[i]);
        }
        SendMessage(CbParity, CB_SETCURSEL, (WPARAM)0, (LPARAM)0); // 选择默认项
        //开关
        BtnPort = GetDlgItem(hwndDlg, IDC_BUTTON_PORT);
        SendMessage(BtnPort, BM_SETCHECK, BST_CHECKED, (LPARAM)0);
        SetWindowText(BtnPort, BtnPortStateText[BtnPortState]);
        ret = (INT_PTR)TRUE;

        //提示信息
        TcMsg = GetDlgItem(hwndDlg, IDC_STATIC_MSG);
        BtnClear = GetDlgItem(hwndDlg, IDC_BUTTON_CLEAR);
        BtnSend = GetDlgItem(hwndDlg, IDC_BUTTON_SEND);
        EditReceive = GetDlgItem(hwndDlg, IDC_EDIT_RECEIVE);
        SendMessage(EditReceive, EM_SETREADONLY, TRUE, 0);
        EditSend = GetDlgItem(hwndDlg, IDC_EDIT_SEND);
        break;
    }

    case WM_DEVICECHANGE:
        if (wParam == DBT_DEVICEREMOVECOMPLETE) {
            PDEV_BROADCAST_HDR lpdb = (PDEV_BROADCAST_HDR)lParam;
            if (lpdb->dbch_devicetype == DBT_DEVTYP_PORT) {
                PDEV_BROADCAST_PORT pdbp = (PDEV_BROADCAST_PORT)lpdb;
                TCHAR* portName = pdbp->dbcp_name;
                // 断开的串口是打开的串口
                if (wcsncmp(portName, SerialportName, sizeof(SerialportName) / sizeof(SerialportName[0])) == 0)
                {
                    //提示串口断开
                    GetLocalTime(&st);
                    swprintf_s(FormattedString, sizeof(FormattedString) / sizeof(FormattedString[0]), L"%04d-%02d-%02d %02d:%02d:%02d>%ls断开",
                        st.wYear, st.wMonth, st.wDay,st.wHour, st.wMinute, st.wSecond, portName);
                    wprintf(L"%ls断开\r\n", portName);
                    SetWindowText(TcMsg, FormattedString);

                    SendMessage(CbSerialPort, CB_RESETCONTENT, 0, 0);//清空端口号
                    SerialportName[0] = '\0'; //清空打开串口字符串
                    //将按钮设置为打开串口
                    BtnPortState = 0;
                    SendMessage(BtnPort, WM_SETTEXT, 0, (LPARAM)BtnPortStateText[BtnPortState]);
                    // 使能下拉框选择
                    EnableWindow(CbSerialPort, TRUE);
                    EnableWindow(CbBaudRate, TRUE);
                    EnableWindow(CbByteSize, TRUE);
                    EnableWindow(CbStopBits, TRUE);
                    EnableWindow(CbParity, TRUE);
                    EnableWindow(BtnSend, FALSE);
                }
            }
        }
        ret = (INT_PTR)TRUE;
        break;

    case WM_COMMAND:
    {
        // 低字节指定了组合框的控件标识符，高字节指定了通知消息
        int wmId = LOWORD(wParam);
        int wmEvent = HIWORD(wParam);
        // 检查是否是组合框的选择改变消息
        if (wmEvent == CBN_DROPDOWN) {
            if (wmId == IDC_SERIALPORT) { // 替换为组合框的ID
                SendMessage(CbSerialPort, CB_RESETCONTENT, (WPARAM)0, (LPARAM)0);
                FindValidSerialPort(CbSerialPort);
            }
        }
        if (wmEvent == BN_CLICKED) {
            if (wmId == IDC_BUTTON_PORT) {
                // 调用ComboBox_GetText函数
                // 第一个参数是ComboBox控件的句柄
                // 第二个参数是接收文本的缓冲区
                // 第三个参数是缓冲区的大小（字符数）
                if (ComboBox_GetText(CbSerialPort, SerialportName, sizeof(SerialportName) / sizeof(SerialportName[0])) > 0)
                {
                    // ComboBox_GetText函数成功，SerialportName现在包含选中项的文本
                    if (BtnPortState == 0) {
                        // 按钮被选中，串口打开
                        BtnPortState = 1;
                        // 禁用下拉框选择
                        EnableWindow(CbSerialPort, FALSE);
                        EnableWindow(CbBaudRate, FALSE);
                        EnableWindow(CbByteSize, FALSE);
                        EnableWindow(CbStopBits, FALSE);
                        EnableWindow(CbParity, FALSE);
                        EnableWindow(BtnSend, TRUE);

                        // 创建串口句柄
                        hComPort = CreateFile(
                            SerialportName,
                            GENERIC_READ | GENERIC_WRITE,
                            0,
                            NULL,
                            OPEN_EXISTING,
                            FILE_ATTRIBUTE_NORMAL,
                            NULL
                        );

                        if (hComPort == INVALID_HANDLE_VALUE) {
                            wprintf(L"Error opening port %s. Error code: %d", SerialportName, GetLastError());
                            ret = (INT_PTR)FALSE;
                            break;
                        }

                        ComboBox_GetText(CbBaudRate, SerialBaudRate, sizeof(SerialBaudRate) / sizeof(SerialBaudRate[0]));
                        BaudRate = (DWORD)_wtoi(SerialBaudRate);
                        ComboBox_GetText(CbByteSize, SerialByteSize, sizeof(SerialByteSize) / sizeof(SerialByteSize[0]));
                        ByteSize = (BYTE)_wtoi(SerialByteSize);

                        // 配置串口参数，例如波特率、数据位等
                        DCB dcbSerialParams = { 0 };
                        COMMTIMEOUTS timeouts = { 0 };
                        dcbSerialParams.DCBlength = sizeof(dcbSerialParams);
                        if (!GetCommState(hComPort, &dcbSerialParams)) {
                            wprintf(L"Error getting serial port state. Error code: %d", GetLastError());
                            CloseHandle(hComPort);
                            ret = (INT_PTR)FALSE;
                            break;
                        }
                        dcbSerialParams.BaudRate = BaudRate;
                        dcbSerialParams.ByteSize = ByteSize;
                        dcbSerialParams.StopBits = ComboBox_GetCurSel(CbStopBits);
                        dcbSerialParams.Parity = ComboBox_GetCurSel(CbParity);
                        if (!SetCommState(hComPort, &dcbSerialParams)) {
                            wprintf(L"Error setting serial port state. Error code: %d", GetLastError());
                            CloseHandle(hComPort);
                            ret = (INT_PTR)FALSE;
                            break;
                        }

                        timeouts.ReadIntervalTimeout = MAXDWORD;
                        timeouts.ReadTotalTimeoutConstant = 0;
                        timeouts.ReadTotalTimeoutMultiplier = 0;
                        timeouts.WriteTotalTimeoutConstant = 0;
                        timeouts.WriteTotalTimeoutMultiplier = 0;

                        if (!SetCommTimeouts(hComPort, &timeouts)) {
                            CloseHandle(hComPort);
                            wprintf(L"Error setting serial port timeouts. Error code: %d", GetLastError());
                            ret = (INT_PTR)FALSE;
                            break;
                        }

                        //提示串口打开成功
                        GetLocalTime(&st);
                        swprintf_s(FormattedString, sizeof(FormattedString) / sizeof(FormattedString[0]), L"%04d-%02d-%02d %02d:%02d:%02d>打开%ls成功",
                            st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond, SerialportName);
                        wprintf(L"%ls\r\n", FormattedString);
                        SetWindowText(TcMsg, FormattedString);
                    }
                    else {
                        // 按钮未被选中，串口关闭
                        BtnPortState = 0;
                        // 使能下拉框选择
                        EnableWindow(CbSerialPort, TRUE);
                        EnableWindow(CbBaudRate, TRUE);
                        EnableWindow(CbByteSize, TRUE);
                        EnableWindow(CbStopBits, TRUE);
                        EnableWindow(CbParity, TRUE);
                        EnableWindow(BtnSend, FALSE);

                        CloseHandle(hComPort);
                        //提示串口关闭成功
                        GetLocalTime(&st);
                        swprintf_s(FormattedString, sizeof(FormattedString) / sizeof(FormattedString[0]), L"%04d-%02d-%02d %02d:%02d:%02d>关闭%ls成功",
                            st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond, SerialportName);
                        wprintf(L"%ls\r\n", FormattedString);
                        SetWindowText(TcMsg, FormattedString);
                    }
                    SendMessage(BtnPort, WM_SETTEXT, 0, (LPARAM)BtnPortStateText[BtnPortState]);
                }
                else
                {
                    // 处理错误情况，例如，ComboBox可能没有选中任何项
                    GetLocalTime(&st);
                    swprintf_s(FormattedString, sizeof(FormattedString) / sizeof(FormattedString[0]), L"%04d-%02d-%02d %02d:%02d:%02d>没有可用串口", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
                    wprintf(L"%ls\r\n", FormattedString);
                    SetWindowText(TcMsg, FormattedString);
                }
            }

            if (wmId == IDC_BUTTON_SEND) {
                LPTSTR lpszText = NULL;
                int nLength = (int)SendMessage(EditSend, WM_GETTEXTLENGTH, 0, 0);
                if (nLength > 0) {
                    // 为文本分配内存，加1是为了包含字符串结束符'\0'
                    lpszText = (LPTSTR)malloc((nLength + 1) * sizeof(TCHAR));
                    if (lpszText != NULL) {
                        // 获取文本
                        SendMessage(EditSend, WM_GETTEXT, nLength + 1, (LPARAM)lpszText);
                        wprintf(L"%ls\r\n", lpszText);
                    }
                    SendData(lpszText, (nLength + 1) * sizeof(TCHAR));
                    free(lpszText);
                }
            }
            if (wmId == IDC_BUTTON_CLEAR) {
                // 清空多行Edit控件的内容并重置滚动位置
                SendMessage(EditReceive, EM_SETSEL, 0, -1); // 选择Edit控件中的所有文本
                SendMessage(EditReceive, EM_REPLACESEL, 0, (LPARAM)""); // 用空字符串替换选中的文本
            }
        }
        ret = (INT_PTR)TRUE;
        break;
    }
    case WM_TIMER:
        // 定时读取串口数据
        ReadDataAndUpdateEdit();
        break;
    case WM_CLOSE:
        EndDialog(hwndDlg, 0);
        ret = (INT_PTR)TRUE;
        break;

    case WM_DESTROY:
        // 清理工作
        KillTimer(hwndDlg, 1);
        PostQuitMessage(0);
        ret = (INT_PTR)TRUE;
        break;

    default:
        ret = (INT_PTR)FALSE;
        break;
    }
    return ret;
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);
    setlocale(LC_ALL, "zh_CN.UTF-8");
    DialogBox(hInstance, MAKEINTRESOURCE(IDD_SERIALPORT), NULL, DialogProc);

    MSG msg;

    // 主消息循环:
    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return (int) msg.wParam;
}
