---
name: vba-optimizer
description: VBA code optimization advisor. Analyzes VBA code and provides specific optimization suggestions based on best practices. Use when reviewing or optimizing Excel macros.
---

# VBA Optimizer

## Overview

This skill provides detailed VBA code optimization advice based on established best practices and performance patterns.

## When to Use This Skill

Use this skill when:
- Reviewing existing VBA code for optimization opportunities
- Analyzing macro performance issues
- Implementing best practices in VBA development
- User explicitly requests VBA optimization analysis

## Analysis Framework

When analyzing VBA code, consider the following optimization categories:

### 1. Screen Update and Calculation Control
- Check for `Application.ScreenUpdating = False`
- Verify `Application.Calculation = xlCalculationManual` for heavy calculations
- Ensure proper restoration in error handlers

### 2. Object References
- Identify unnecessary `Activate` or `Select` calls
- Recommend direct object references instead
- Check for With statements usage

### 3. Array Operations
- Look for cell-by-cell operations that could use arrays
- Suggest bulk read/write patterns
- Identify opportunities for array processing

### 4. Loop Optimization
- Check for unnecessary Range operations inside loops
- Recommend Dictionary objects for lookups
- Suggest batch operations where applicable

### 5. Error Handling
- Verify presence of error handlers
- Check for proper cleanup in error cases
- Ensure Application properties are restored

## Output Format

Provide optimization suggestions in this structure:

```
### Issue: [Brief description]
**Current Code:**
[Problematic code snippet]

**Optimized Code:**
[Improved version]

**Explanation:**
[Why this is better, expected performance impact]
```

## Key Principles

1. **Specific over generic**: Reference actual code patterns found
2. **Quantify impact**: Estimate performance improvements when possible
3. **Maintain readability**: Don't sacrifice code clarity for micro-optimizations
4. **Error safety**: Always consider error handling implications

## Reference Knowledge Base

Refer to project documentation when available:
- `/docs/excel-knowledge/patterns/VBA_OPTIMIZATION_PATTERNS.md`
- `/docs/excel-knowledge/failures/` - Learn from past mistakes
- Project-specific coding standards in `CLAUDE.md`

## Language

- Provide explanations in Japanese
- Use technical terms accurately
- Reference specific Excel/VBA API methods correctly
