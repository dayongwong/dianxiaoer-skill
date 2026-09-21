#!/usr/bin/env python3
import os
import sys
import json
import platform
import subprocess
from pathlib import Path

def print_step(msg):
    print(f"\n=====================================")
    print(f"👉 {msg}")
    print(f"=====================================\n")

def check_python_version():
    if sys.version_info < (3, 8):
        print("❌ 错误：店小二插件需要 Python 3.8 或以上版本。")
        input("按回车键退出...")
        sys.exit(1)

def install_dependencies():
    print_step("正在安装运行环境依赖...")
    req_file = Path(__file__).parent / "requirements.txt"
    if not req_file.exists():
        print("未找到 requirements.txt，跳过依赖安装。")
        return
    
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", str(req_file)])
        print("✅ 依赖安装成功！")
    except Exception as e:
        print(f"❌ 依赖安装失败：{e}")
        input("按回车键退出...")
        sys.exit(1)

def configure_mcp(agent_choice):
    print_step("正在为您配置 AI Agent 插件...")
    
    home_dir = Path.home()
    config_path = None
    
    if agent_choice == '1': # Antigravity (AGY)
        config_path = home_dir / ".gemini" / "config" / "mcp_config.json"
    elif agent_choice == '2': # Claude Desktop
        if platform.system() == "Darwin":
            config_path = home_dir / "Library" / "Application Support" / "Claude" / "claude_desktop_config.json"
        elif platform.system() == "Windows":
            config_path = home_dir / "AppData" / "Roaming" / "Claude" / "claude_desktop_config.json"
    
    if not config_path:
        print("❌ 暂不支持自动配置该客户端，请参考文档手动配置。")
        return
    
    # 准备店小二的 MCP 配置节点
    server_script = str((Path(__file__).parent / "mcp_server" / "server.py").resolve())
    dianxiaoer_config = {
        "command": sys.executable,
        "args": [server_script]
    }
    
    # 读取并修改配置
    try:
        config_path.parent.mkdir(parents=True, exist_ok=True)
        config_data = {}
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                    if content:
                        config_data = json.loads(content)
            except Exception as e:
                print(f"⚠️ 读取原有配置失败，将创建新配置：{e}")
        
        # 针对不同客户端的 JSON 结构处理
        if "mcpServers" not in config_data:
            config_data["mcpServers"] = {}
            
        config_data["mcpServers"]["dianxiaoer"] = dianxiaoer_config
        
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, ensure_ascii=False, indent=2)
            
        print(f"✅ 成功将插件挂载到配置文件：{config_path}")
        print("🎉 恭喜！店小二一键选品插件安装完成！")
        print("➡️  现在请【重启您的 AI 客户端】，然后在聊天框输入：“帮我选一款爆款充电宝” 即可体验！")
        
    except Exception as e:
        print(f"❌ 写入配置失败：{e}")

def main():
    print("==================================================")
    print("        欢迎使用【店小二跨境选品】一键安装程序        ")
    print("==================================================")
    
    check_python_version()
    install_dependencies()
    
    print_step("请选择您正在使用的 AI 客户端 (输入数字并回车)：")
    print("1. Antigravity (AGY)")
    print("2. Claude Desktop")
    print("3. 其他 (跳过自动配置，我将手动配置)")
    
    choice = input("您的选择 [1/2/3]: ").strip()
    
    if choice in ['1', '2']:
        configure_mcp(choice)
    else:
        print("已跳过自动配置，请查阅 README.md 手动挂载。")
        
    input("\n安装流程结束，按回车键退出...")

if __name__ == "__main__":
    main()
