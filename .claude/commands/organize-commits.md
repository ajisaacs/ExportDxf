# Organize Changes into Logical Commits

Analyze all current git changes and organize them into logical, atomic commits. Follow these steps:

1. **Analyze Changes**: Run git status and git diff to see all modified and untracked files
2. **Review Content**: Examine the actual changes in each file to understand what was modified
3. **Group Logically**: Group changes by:
   - Feature or bug fix
   - Service or component
   - Related functionality
   - UI changes vs business logic vs API changes
4. **Create Commits**: For each logical group:
   - Stage only the relevant files
   - Create a descriptive commit message following conventional commit format
   - Use prefixes like feat:, fix:, refactor:, chore:, docs:, etc.
5. **Verify**: After all commits, show git log to confirm all changes were committed

Important guidelines:
- Keep commits atomic (one logical change per commit)
- Write clear, descriptive commit messages
- Don't mix unrelated changes in the same commit
- Follow the existing commit message style in the repository
- Include the Claude Code attribution at the end of each commit message
