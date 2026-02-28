#![windows_subsystem = "windows"]

use clipboard::ClipboardProvider;
use clipboard::ClipboardContext;
use std::env;
use std::fs::File;
use std::io::Write;
use std::path::Path;
use winapi::um::winuser::{MessageBeep, MB_ICONASTERISK, GetActiveWindow, MessageBoxW, MB_OK, MB_ICONINFORMATION, MB_ICONERROR};
use winapi::um::commdlg::{GetSaveFileNameW, OPENFILENAMEW};
use winapi::um::processthreadsapi::{GetCurrentProcess, OpenProcessToken};
use winapi::um::securitybaseapi::GetTokenInformation;
use winapi::um::winnt::{TOKEN_READ, TOKEN_QUERY, TOKEN_ELEVATION};
use winapi::um::handleapi::CloseHandle;
use winreg::RegKey;
use winreg::enums::*;
use std::ffi::OsString;
use std::os::windows::ffi::{OsStringExt, OsStrExt};

// 检查是否具有管理员权限
fn is_admin() -> bool {
    use std::ptr;
    use std::mem;
    
    let mut token = ptr::null_mut();
    if unsafe { OpenProcessToken(GetCurrentProcess(), TOKEN_READ | TOKEN_QUERY, &mut token) } == 0 {
        return false;
    }
    
    let mut elevation = TOKEN_ELEVATION {
        TokenIsElevated: 0
    };
    let mut size = mem::size_of::<TOKEN_ELEVATION>() as u32;
    
    const TokenElevation: u32 = 20;
    
    let result = unsafe {
        GetTokenInformation(
            token,
            TokenElevation,
            &mut elevation as *mut _ as *mut _, 
            size,
            &mut size
        )
    };
    
    unsafe {
        CloseHandle(token);
    }
    
    result != 0 && elevation.TokenIsElevated != 0
}

fn main() {
    let args: Vec<String> = env::args().collect();
    
    // 检查是否有命令行参数
    if args.len() > 1 {
        // 有参数，说明是从右键菜单调用
        let path = &args[1];
        generate_readme(path);
    } else {
        // 无参数，注册右键菜单
        // 检查是否具有管理员权限
        if !is_admin() {
            // 显示权限不足的提示
            let title: Vec<u16> = OsString::from("权限不足").encode_wide().chain(Some(0)).collect::<Vec<_>>();
            let message: Vec<u16> = OsString::from("需要管理员权限才能注册或取消注册右键菜单！\n请以管理员身份运行该程序。").encode_wide().chain(Some(0)).collect::<Vec<_>>();
            
            unsafe {
                MessageBoxW(
                    GetActiveWindow(),
                    message.as_ptr(),
                    title.as_ptr(),
                    MB_OK | MB_ICONERROR
                );
            }
            return;
        }
        
        register_context_menu();
    }
}

// 注册或取消注册右键菜单
fn register_context_menu() {
    // 打开注册表中的文件夹右键菜单位置
    let hkcr = RegKey::predef(HKEY_CLASSES_ROOT);
    
    // 检查是否已经注册过
    if hkcr.open_subkey("Directory\\shell\\ReadmeGen4").is_ok() {
        // 已经注册过，删除右键菜单项
        hkcr.delete_subkey_all("Directory\\shell\\ReadmeGen4").unwrap();
        
        // 显示取消注册成功的对话框
        let title: Vec<u16> = OsString::from("ReadmeGen4 取消注册成功").encode_wide().chain(Some(0)).collect::<Vec<_>>();
        let message: Vec<u16> = OsString::from("ReadmeGen4已经从目录对象中移除！\n右键菜单中的 '生成readme.txt' 选项已删除。").encode_wide().chain(Some(0)).collect::<Vec<_>>();
        
        unsafe {
            MessageBoxW(
                GetActiveWindow(),
                message.as_ptr(),
                title.as_ptr(),
                MB_OK | MB_ICONINFORMATION
            );
        }
    } else {
        // 未注册，创建右键菜单项
        let exe_path = env::current_exe().unwrap();
        let exe_path_str = exe_path.to_str().unwrap();
        
        let (folder_key, _) = hkcr.create_subkey("Directory\\shell\\ReadmeGen4").unwrap();
        
        // 设置菜单名称
        folder_key.set_value("", &"生成readme.txt").unwrap();
        
        // 创建命令子键
        let (command_key, _) = folder_key.create_subkey("command").unwrap();
        
        // 设置命令，带参数 %V（当前文件夹路径）
        let command = format!("\"{}\" \"%V\"", exe_path_str);
        command_key.set_value("", &command).unwrap();
        
        // 显示注册成功的对话框
        let title: Vec<u16> = OsString::from("ReadmeGen4 注册成功").encode_wide().chain(Some(0)).collect::<Vec<_>>();
        let message: Vec<u16> = OsString::from("ReadmeGen4已经关联到目录对象！\n在资源管理器中，对目录按右键使用该功能。").encode_wide().chain(Some(0)).collect::<Vec<_>>();
        
        unsafe {
            MessageBoxW(
                GetActiveWindow(),
                message.as_ptr(),
                title.as_ptr(),
                MB_OK | MB_ICONINFORMATION
            );
        }
    }
}

// 生成 readme.txt 文件
fn generate_readme(path: &str) {
    // 获取剪贴板内容
    let mut ctx: ClipboardContext = ClipboardProvider::new().unwrap();
    let clipboard_content = ctx.get_contents().unwrap_or_else(|_| "".to_string());
    
    // 构建默认的 readme.txt 文件路径
    let mut file_path = Path::new(path).join("readme.txt");
    
    // 检查文件是否存在
    if file_path.exists() {
        // 文件存在，显示保存对话框
        if let Some(new_path) = show_save_dialog(path) {
            file_path = new_path;
        } else {
            // 用户取消了保存
            println!("保存操作已取消");
            return;
        }
    }
    
    // 写入文件（使用 UTF-8 编码）
    match File::create(&file_path) {
        Ok(mut file) => {
            if let Err(e) = write!(file, "{}", clipboard_content) {
                println!("写入文件失败: {}", e);
            } else {
                println!("文件生成成功: {:?}", file_path);
                // 鸣叫主板喇叭
                unsafe {
                    MessageBeep(MB_ICONASTERISK);
                }
            }
        },
        Err(e) => {
            println!("创建文件失败: {}", e);
        }
    }
}

// 显示保存对话框
fn show_save_dialog(directory: &str) -> Option<std::path::PathBuf> {
    use std::mem;
    
    // 构建默认文件名
    let default_file = "readme.txt";
    
    // 准备文件名缓冲区
    let mut file_name: [u16; 260] = [0; 260];
    for (i, c) in default_file.encode_utf16().enumerate() {
        if i < file_name.len() {
            file_name[i] = c;
        }
    }
    
    // 准备 OPENFILENAMEW 结构体
    let filter = "Text Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0\0".encode_utf16().collect::<Vec<_>>();
    let title = "保存文件".encode_utf16().collect::<Vec<_>>();
    
    let mut ofn: OPENFILENAMEW = unsafe {
        mem::zeroed()
    };
    ofn.lStructSize = mem::size_of::<OPENFILENAMEW>() as u32;
    ofn.hwndOwner = unsafe { GetActiveWindow() };
    ofn.lpstrFile = file_name.as_mut_ptr();
    ofn.nMaxFile = file_name.len() as u32;
    ofn.lpstrFilter = filter.as_ptr();
    ofn.nFilterIndex = 1;
    ofn.lpstrTitle = title.as_ptr();
    ofn.Flags = 0x00000002; // OFN_OVERWRITEPROMPT
    
    // 显示保存对话框
    let result = unsafe {
        GetSaveFileNameW(&mut ofn)
    };
    
    if result != 0 {
        // 解析文件名
        let file_name_str = OsString::from_wide(&file_name[..file_name.iter().position(|&c| c == 0).unwrap()]);
        let file_path = Path::new(directory).join(file_name_str);
        Some(file_path)
    } else {
        None
    }
}
