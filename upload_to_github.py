# -*- coding: utf-8 -*-
"""
上传 Streamlit 应用到 GitHub
=============================
使用 GitHub API 直接上传文件，无需安装 git。

使用方法:
  python upload_to_github.py <token>

1. 在 GitHub Settings > Developer settings > Personal access tokens > Tokens (classic)
2. 勾选 repo 权限
3. 运行: python upload_to_github.py ghp_xxxxxxxxxxxx
"""
import sys
import os
from github import Github

# ---- 配置 ----
REPO_NAME = "sales-forecast-bom-decompose"
REPO_DESC = "销售预测捆绑SKU拆解 Streamlit 小程序"
APP_DIR = os.path.dirname(os.path.abspath(__file__))

# 要上传的文件（相对路径）
FILES_TO_UPLOAD = [
    "app.py",
    "decompose.py",
    "requirements.txt",
    "run.bat",
    ".gitignore",
    "README.md",
    "upload_to_github.py",
]

COMMIT_MSG = "初始提交: 销售预测捆绑SKU拆解 Streamlit 应用"


def main():
    if len(sys.argv) < 2:
        print("用法: python upload_to_github.py <GitHub_Token>")
        print()
        print("获取 Token 步骤:")
        print("  1. 打开 https://github.com/settings/tokens")
        print("  2. 点击 'Generate new token (classic)'")
        print("  3. 勾选 'repo' 权限")
        print("  4. 生成后复制 token")
        print("  5. 运行: python upload_to_github.py ghp_xxxxxxxxxxxx")
        sys.exit(1)

    token = sys.argv[1].strip()
    gh = Github(token)

    # 获取当前用户
    try:
        user = gh.get_user()
        print(f"GitHub 用户: {user.login}")
    except Exception as e:
        print(f"Token 验证失败: {e}")
        sys.exit(1)

    # 创建仓库（如果不存在）
    try:
        repo = user.get_repo(REPO_NAME)
        print(f"仓库已存在: {repo.full_name}")
    except:
        print(f"创建新仓库: {REPO_NAME}")
        repo = gh.get_user().create_repo(
            REPO_NAME,
            description=REPO_DESC,
            private=False,
            auto_init=False,
        )
        print(f"  仓库已创建: {repo.html_url}")

    # 上传文件
    print(f"\n上传文件到 {repo.full_name}...")
    uploaded = 0
    skipped = 0

    for rel_path in FILES_TO_UPLOAD:
        file_path = os.path.join(APP_DIR, rel_path)
        if not os.path.exists(file_path):
            print(f"  [跳过] {rel_path} (文件不存在)")
            skipped += 1
            continue

        with open(file_path, "rb") as f:
            content_bytes = f.read()

        # PyGithub 内部会自动 base64 编码，不需要手动编码
        # 检查文件是否已存在
        try:
            existing = repo.get_contents(rel_path)
            # 更新已有文件
            repo.update_file(
                path=rel_path,
                message=f"更新: {rel_path}",
                content=content_bytes,
                sha=existing.sha,
            )
            print(f"  [更新] {rel_path} ({len(content_bytes):,} bytes)")
        except:
            # 创建新文件
            repo.create_file(
                path=rel_path,
                message=COMMIT_MSG if uploaded == 0 else f"添加: {rel_path}",
                content=content_bytes,
            )
            print(f"  [上传] {rel_path} ({len(content_bytes):,} bytes)")

        uploaded += 1

    print(f"\n完成! 上传 {uploaded} 个文件, 跳过 {skipped} 个")
    print(f"仓库地址: {repo.html_url}")
    print(f"克隆命令: git clone {repo.clone_url}")


if __name__ == "__main__":
    main()
