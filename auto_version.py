#!/usr/bin/env python3
"""
Automatic version bumping based on conventional commits.
Analyzes commit history and updates version accordingly.
"""

import subprocess
import re
import sys
from pathlib import Path
from typing import Optional, List


VERSION_FILE = Path("src/_version.py")


def run_git(cmd: str) -> Optional[str]:
    """Run git command and return output."""
    try:
        result = subprocess.run(
            cmd, shell=True, capture_output=True, text=True, check=True
        )
        return result.stdout.strip()
    except subprocess.CalledProcessError:
        return None


def get_current_version() -> str:
    """Get current version from version file."""
    if VERSION_FILE.exists():
        content = VERSION_FILE.read_text()
        match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', content)
        if match:
            return match.group(1)
    return "0.0.0"


def get_last_tag() -> Optional[str]:
    """Get the last git tag."""
    tag = run_git("git describe --tags --abbrev=0 2>nul")
    if tag and tag.startswith('v'):
        return tag[1:]
    return None


def get_commits_since_tag(tag: Optional[str]) -> List[str]:
    """Get commit messages since the last tag."""
    if tag:
        commits = run_git(f'git log v{tag}..HEAD --pretty=format:"%s"')
    else:
        commits = run_git('git log --pretty=format:"%s"')
    
    return commits.split('\n') if commits else []


def determine_bump_type(commits: List[str]) -> Optional[str]:
    """Determine version bump type from commits."""
    if not commits:
        return None
    
    has_major = False
    has_minor = False
    has_patch = False
    
    for commit in commits:
        commit_lower = commit.lower()
        
        # Check for breaking changes (major)
        if 'breaking change' in commit_lower or '!' in commit[:20]:
            has_major = True
            break
        
        # Check for features (minor)
        if commit_lower.startswith('feat'):
            has_minor = True
        
        # Check for fixes (patch)
        if commit_lower.startswith(('fix', 'perf', 'refactor')):
            has_patch = True
    
    if has_major:
        return 'major'
    elif has_minor:
        return 'minor'
    elif has_patch:
        return 'patch'
    
    return None


def bump_version(current: str, bump_type: str) -> str:
    """Calculate new version."""
    match = re.match(r'(\d+)\.(\d+)\.(\d+)', current)
    if not match:
        return "1.0.0"
    
    major, minor, patch = map(int, match.groups())
    
    if bump_type == 'major':
        return f"{major + 1}.0.0"
    elif bump_type == 'minor':
        return f"{major}.{minor + 1}.0"
    else:  # patch
        return f"{major}.{minor}.{patch + 1}"


def update_version_file(version: str) -> None:
    """Update version in _version.py."""
    content = f'''"""Version information for Motor Report Application."""
__version__ = "{version}"
'''
    VERSION_FILE.write_text(content, encoding='utf-8')
    print(f"✓ Updated {VERSION_FILE} to {version}")


def create_tag(version: str) -> None:
    """Create git tag for version."""
    tag = f"v{version}"
    
    run_git(f'git add {VERSION_FILE}')
    run_git(f'git commit -m "chore(release): bump version to {version}"')
    run_git(f'git tag -a {tag} -m "Release {version}"')
    
    print(f"✓ Created tag {tag}")
    print(f"\n💡 Push with: git push && git push origin {tag}")


def main():
    """Main entry point."""
    dry_run = '--dry-run' in sys.argv
    
    print("🔍 Analyzing commit history...\n")
    
    current_version = get_current_version()
    last_tag = get_last_tag()
    commits = get_commits_since_tag(last_tag)
    
    print(f"📌 Current version: {current_version}")
    print(f"📝 New commits since last release: {len(commits)}")
    
    if not commits:
        print("✓ No new commits, no version bump needed")
        return 0
    
    bump_type = determine_bump_type(commits)
    
    if not bump_type:
        print("✓ No version-bumping commits found (use feat/fix/refactor)")
        return 0
    
    new_version = bump_version(current_version, bump_type)
    
    print(f"\n📈 Version bump type: {bump_type.upper()}")
    print(f"🎯 {current_version} → {new_version}")
    
    if dry_run:
        print(f"\n🔍 [DRY RUN] Would update to {new_version}")
        return 0
    
    update_version_file(new_version)
    create_tag(new_version)
    
    print(f"\n✅ Version bumped successfully to {new_version}")
    
    return 0


if __name__ == "__main__":
    sys.exit(main())
