"""
HR Slide Engine — Installation script.

Copies the skill to ~/.claude/skills/pptx-master-rh/
Copies slide_engine module to a persistent location.
Replaces {MODULE_PATH} in the installed SKILL.md.
"""

import os
import sys
import shutil
import platform


def get_paths():
    """Get platform-specific installation paths."""
    if platform.system() == "Windows":
        home = os.environ.get("USERPROFILE", os.path.expanduser("~"))
        module_dir = os.path.join(home, ".local", "lib", "hr-slide-engine")
    else:
        home = os.path.expanduser("~")
        module_dir = os.path.join(home, ".local", "lib", "hr-slide-engine")

    skill_dir = os.path.join(home, ".claude", "skills", "pptx-master-rh")
    return home, module_dir, skill_dir


def copy_tree(src, dst):
    """Copy directory tree, overwriting existing files."""
    if os.path.exists(dst):
        shutil.rmtree(dst)
    shutil.copytree(src, dst)


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    home, module_dir, skill_dir = get_paths()

    print("=" * 50)
    print("  HR Slide Engine — Installation")
    print("=" * 50)
    print()

    # 1. Copy slide_engine module
    src_module = os.path.join(script_dir, "slide_engine")
    if not os.path.exists(src_module):
        print(f"ERROR: slide_engine/ not found in {script_dir}")
        sys.exit(1)

    os.makedirs(os.path.dirname(module_dir), exist_ok=True)
    copy_tree(src_module, module_dir)
    print(f"[OK] Module installed: {module_dir}")

    # 2. Copy skill files
    src_skill = os.path.join(script_dir, "skill")
    if not os.path.exists(src_skill):
        print(f"ERROR: skill/ not found in {script_dir}")
        sys.exit(1)

    os.makedirs(os.path.dirname(skill_dir), exist_ok=True)
    copy_tree(src_skill, skill_dir)
    print(f"[OK] Skill installed: {skill_dir}")

    # 3. Replace {MODULE_PATH} in SKILL.md
    skill_md = os.path.join(skill_dir, "SKILL.md")
    with open(skill_md, "r", encoding="utf-8") as f:
        content = f.read()

    # Use the parent directory of slide_engine so imports work
    module_parent = os.path.dirname(module_dir)
    content = content.replace("{MODULE_PATH}", module_parent)

    with open(skill_md, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"[OK] Module path set: {module_parent}")

    # 4. Check python-pptx
    print()
    try:
        import pptx
        print(f"[OK] python-pptx {pptx.__version__} is installed")
    except ImportError:
        print("[!!] python-pptx not found. Install it:")
        print("     pip install python-pptx")

    print()
    print("=" * 50)
    print("  Installation complete!")
    print()
    print("  Usage in Claude Code:")
    print("    /pptx-master-rh <sujet ou texte brut>")
    print()
    print("  Example:")
    print('    /pptx-master-rh La GPEC comme levier stratégique')
    print("=" * 50)


if __name__ == "__main__":
    main()
