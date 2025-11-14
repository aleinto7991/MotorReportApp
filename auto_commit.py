#!/usr/bin/env python3
"""
Fully automated commit message generator using local analysis.
Analyzes git diff and generates conventional commit messages without external AI.
"""

import subprocess
import re
import sys
from pathlib import Path
from typing import List, Dict, Tuple, Optional


class SmartCommitGenerator:
    """Generate smart commit messages by analyzing code changes."""
    
    # Keywords that indicate different types of changes
    FEATURE_KEYWORDS = ['add', 'new', 'create', 'implement', 'feature']
    FIX_KEYWORDS = ['fix', 'bug', 'error', 'correct', 'patch', 'resolve']
    REFACTOR_KEYWORDS = ['refactor', 'restructure', 'reorganize', 'cleanup', 'optimize']
    DOCS_KEYWORDS = ['doc', 'comment', 'readme', 'documentation']
    PERF_KEYWORDS = ['performance', 'speed', 'optimize', 'fast', 'cache']
    TEST_KEYWORDS = ['test', 'spec', 'pytest']
    
    # File patterns to scope mapping
    FILE_SCOPES = {
        'reports': ['reports/', 'builders/', 'sap_sheet', 'comparison_sheet'],
        'ui': ['ui/', 'gui/', 'main_gui', 'components/'],
        'charts': ['chart', 'noise_chart', 'graph'],
        'services': ['services/', 'registry', 'loader'],
        'config': ['config/', 'settings', 'directory_config'],
        'utils': ['utils/', 'helpers'],
        'core': ['main.py', 'models.py'],
    }
    
    def __init__(self):
        self.staged_files = []
        self.diff_content = ""
        
    def run_git(self, cmd: str) -> str:
        """Run git command and return output."""
        try:
            result = subprocess.run(
                cmd, shell=True, capture_output=True, text=True, 
                check=True, encoding='utf-8', errors='ignore'
            )
            return result.stdout.strip() if result.stdout else ""
        except subprocess.CalledProcessError as e:
            return ""
        except Exception as e:
            print(f"Warning: Git command failed: {e}")
            return ""
    
    def get_staged_files(self) -> List[str]:
        """Get list of staged files."""
        output = self.run_git("git diff --cached --name-only")
        if not output:
            return []
        return [f for f in output.split('\n') if f.strip()]
    
    def get_diff(self) -> str:
        """Get the actual diff content."""
        diff = self.run_git("git diff --cached")
        return diff if diff else ""
    
    def determine_scope(self, files: List[str]) -> Optional[str]:
        """Determine the scope based on modified files."""
        for scope, patterns in self.FILE_SCOPES.items():
            for file in files:
                file_lower = file.lower()
                for pattern in patterns:
                    if pattern in file_lower:
                        return scope
        return None
    
    def analyze_changes(self, diff: str) -> Dict[str, int]:
        """Analyze the diff to understand what changed."""
        stats = {
            'additions': 0,
            'deletions': 0,
            'feature_indicators': 0,
            'fix_indicators': 0,
            'refactor_indicators': 0,
            'docs_indicators': 0,
            'perf_indicators': 0,
            'test_indicators': 0,
        }
        
        for line in diff.split('\n'):
            if line.startswith('+') and not line.startswith('+++'):
                stats['additions'] += 1
                line_lower = line.lower()
                
                # Count keyword indicators
                if any(kw in line_lower for kw in self.FEATURE_KEYWORDS):
                    stats['feature_indicators'] += 1
                if any(kw in line_lower for kw in self.FIX_KEYWORDS):
                    stats['fix_indicators'] += 1
                if any(kw in line_lower for kw in self.REFACTOR_KEYWORDS):
                    stats['refactor_indicators'] += 1
                if any(kw in line_lower for kw in self.DOCS_KEYWORDS):
                    stats['docs_indicators'] += 1
                if any(kw in line_lower for kw in self.PERF_KEYWORDS):
                    stats['perf_indicators'] += 1
                if any(kw in line_lower for kw in self.TEST_KEYWORDS):
                    stats['test_indicators'] += 1
                    
            elif line.startswith('-') and not line.startswith('---'):
                stats['deletions'] += 1
        
        return stats
    
    def determine_type(self, stats: Dict[str, int], files: List[str]) -> str:
        """Determine commit type based on analysis."""
        # Check file patterns first
        file_patterns = ' '.join(files).lower()
        
        if any(f.startswith('test') for f in files):
            return 'test'
        if 'readme' in file_patterns or 'doc' in file_patterns:
            return 'docs'
        
        # Check content patterns
        total_indicators = (
            stats['feature_indicators'] + 
            stats['fix_indicators'] + 
            stats['refactor_indicators']
        )
        
        if total_indicators == 0:
            # No clear indicators, use heuristics
            if stats['additions'] > stats['deletions'] * 2:
                return 'feat'
            elif stats['deletions'] > stats['additions']:
                return 'refactor'
            else:
                return 'chore'
        
        # Determine based on highest indicator count
        if stats['fix_indicators'] > 0:
            return 'fix'
        elif stats['feature_indicators'] > 0:
            return 'feat'
        elif stats['perf_indicators'] > 0:
            return 'perf'
        elif stats['refactor_indicators'] > 0:
            return 'refactor'
        elif stats['test_indicators'] > 0:
            return 'test'
        else:
            return 'chore'
    
    def extract_key_changes(self, diff: str, files: List[str]) -> List[str]:
        """Extract key changes from the diff."""
        changes = []
        
        # Analyze file-level changes
        for file in files:
            if 'delete' in file or file.endswith('.pyc'):
                continue
            
            filename = Path(file).name
            
            # Look for function/class definitions
            function_pattern = r'^\+\s*def\s+(\w+)'
            class_pattern = r'^\+\s*class\s+(\w+)'
            
            for line in diff.split('\n'):
                if file in line or filename in line:
                    # Found relevant section
                    func_match = re.search(function_pattern, line)
                    class_match = re.search(class_pattern, line)
                    
                    if func_match:
                        changes.append(f"Add {func_match.group(1)} function")
                    elif class_match:
                        changes.append(f"Add {class_match.group(1)} class")
        
        # Add generic changes based on file types
        if not changes:
            for file in files[:3]:  # Top 3 files
                filename = Path(file).stem
                changes.append(f"Update {filename}")
        
        return changes[:3]  # Limit to 3 bullet points
    
    def generate_commit_message(self) -> str:
        """Generate a conventional commit message."""
        self.staged_files = self.get_staged_files()
        
        if not self.staged_files:
            return "chore: update files"
        
        self.diff_content = self.get_diff()
        
        # Analyze
        stats = self.analyze_changes(self.diff_content)
        commit_type = self.determine_type(stats, self.staged_files)
        scope = self.determine_scope(self.staged_files)
        changes = self.extract_key_changes(self.diff_content, self.staged_files)
        
        # Build commit message
        scope_str = f"({scope})" if scope else ""
        
        # Generate subject line
        if len(self.staged_files) == 1:
            filename = Path(self.staged_files[0]).stem
            subject = f"update {filename}"
        else:
            subject = "update multiple components"
        
        # Refine subject based on type
        if commit_type == 'feat':
            subject = subject.replace('update', 'add')
        elif commit_type == 'fix':
            subject = subject.replace('update', 'fix')
        elif commit_type == 'refactor':
            subject = subject.replace('update', 'refactor')
        
        # Build full message
        message = f"{commit_type}{scope_str}: {subject}\n"
        
        if changes:
            message += "\n"
            for change in changes:
                message += f"- {change}\n"
        
        return message.strip()
    
    def commit(self, message: str) -> bool:
        """Execute the commit."""
        try:
            # Write message to temporary file to handle multiline
            msg_file = Path('.git/COMMIT_EDITMSG_AUTO')
            msg_file.write_text(message, encoding='utf-8')
            
            result = subprocess.run(
                ['git', 'commit', '-F', str(msg_file)],
                capture_output=True,
                text=True
            )
            
            msg_file.unlink()
            
            if result.returncode == 0:
                return True
            else:
                print(f"Commit failed: {result.stderr}")
                return False
                
        except Exception as e:
            print(f"Error committing: {e}")
            return False


def main():
    """Main entry point."""
    dry_run = '--dry-run' in sys.argv
    
    print("🤖 Auto-generating commit message...")
    
    generator = SmartCommitGenerator()
    message = generator.generate_commit_message()
    
    print("\n" + "="*60)
    print("Generated commit message:")
    print("="*60)
    print(message)
    print("="*60)
    
    if dry_run:
        print("\n🔍 [DRY RUN] This is a preview only")
        return 0
    
    # Ask for confirmation
    response = input("\n[C]ommit, [E]dit, or [Q]uit? ").strip().upper()
    
    if response == 'C':
        if generator.commit(message):
            print("✅ Committed successfully!")
            return 0
        else:
            print("❌ Commit failed")
            return 1
    elif response == 'E':
        print("\nEnter your commit message (end with Ctrl+Z on new line):")
        lines = []
        try:
            while True:
                line = input()
                lines.append(line)
        except EOFError:
            pass
        custom_message = '\n'.join(lines).strip()
        if custom_message and generator.commit(custom_message):
            print("✅ Committed successfully!")
            return 0
    else:
        print("❌ Cancelled")
        return 1


if __name__ == "__main__":
    sys.exit(main())
