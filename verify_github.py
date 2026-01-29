#!/usr/bin/env python3
"""
验证GitHub同步状态
"""

import os
import sys
import subprocess
from pathlib import Path

def check_git_status():
    """检查Git状态"""
    print("=" * 60)
    print("GitHub同步验证")
    print("=" * 60)
    
    repo_path = Path(__file__).parent
    
    # 检查Git状态
    print("\n1. 检查Git状态...")
    try:
        result = subprocess.run(
            ["git", "status"],
            cwd=repo_path,
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        if "nothing to commit" in result.stdout:
            print("   [OK] 工作目录干净")
        else:
            print("   [WARN] 有未提交的更改")
            print(f"   输出: {result.stdout[:200]}...")
    except Exception as e:
        print(f"   [ERROR] 错误: {e}")
    
    # 检查远程仓库
    print("\n2. 检查远程仓库...")
    try:
        result = subprocess.run(
            ["git", "remote", "-v"],
            cwd=repo_path,
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        print(f"   远程仓库:")
        for line in result.stdout.strip().split('\n'):
            print(f"     {line}")
    except Exception as e:
        print(f"   [ERROR] 错误: {e}")
    
    # 检查提交历史
    print("\n3. 检查提交历史...")
    try:
        result = subprocess.run(
            ["git", "log", "--oneline", "-5"],
            cwd=repo_path,
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        print("   最近提交:")
        for line in result.stdout.strip().split('\n'):
            print(f"     {line}")
    except Exception as e:
        print(f"   [ERROR] 错误: {e}")
    
    # 检查标签
    print("\n4. 检查版本标签...")
    try:
        result = subprocess.run(
            ["git", "tag", "-l"],
            cwd=repo_path,
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        tags = result.stdout.strip().split('\n')
        if tags and tags[0]:
            print("   标签:")
            for tag in tags:
                print(f"     {tag}")
        else:
            print("   ⚠️  无标签")
    except Exception as e:
        print(f"   ❌ 错误: {e}")
    
    # 检查文件数量
    print("\n5. 检查项目文件...")
    try:
        # 统计各种文件
        py_files = list(repo_path.glob("**/*.py"))
        md_files = list(repo_path.glob("**/*.md"))
        txt_files = list(repo_path.glob("**/*.txt"))
        toml_files = list(repo_path.glob("**/*.toml"))
        
        print(f"   Python文件: {len(py_files)}个")
        print(f"   Markdown文件: {len(md_files)}个")
        print(f"   文本文件: {len(txt_files)}个")
        print(f"   配置文件: {len(toml_files)}个")
        
        # 检查关键文件
        key_files = [
            "office_automation.py",
            "README.md",
            "TEST_REPORT.md",
            "pyproject.toml",
            "setup.py",
            "examples/create_report.py",
            "tests/test_final_integration.py"
        ]
        
        missing_files = []
        for file in key_files:
            if (repo_path / file).exists():
                print(f"   [OK] {file}")
            else:
                missing_files.append(file)
                print(f"   [MISSING] {file} (缺失)")
        
        if missing_files:
            print(f"   [WARN] 缺失关键文件: {len(missing_files)}个")
    
    except Exception as e:
        print(f"   [ERROR] 错误: {e}")
    
    # 运行一个简单测试
    print("\n6. 运行简单测试...")
    try:
        # 导入主模块
        sys.path.insert(0, str(repo_path))
        from office_automation import OfficeAutomation
        
        office = OfficeAutomation()
        info = office.get_info()
        
        print(f"   版本: {info.get('version', 'N/A')}")
        print(f"   模块: {list(info.get('modules', {}).keys())}")
        print("   [OK] 模块导入成功")
        
    except Exception as e:
        print(f"   [ERROR] 模块导入失败: {e}")
    
    print("\n" + "=" * 60)
    print("验证完成!")
    print("=" * 60)
    
    # GitHub链接
    print("\nGitHub链接:")
    print("1. 仓库: https://github.com/chleya/office-automation")
    print("2. 主分支: https://github.com/chleya/office-automation/tree/main")
    print("3. 版本: https://github.com/chleya/office-automation/releases/tag/v1.0.0")
    print("4. 文件: https://github.com/chleya/office-automation/tree/main/examples")
    
    return True

if __name__ == "__main__":
    check_git_status()