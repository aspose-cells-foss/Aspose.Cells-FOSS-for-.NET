# Aspose.Cells FOSS for .NET - Implementation Rules

## Class structure

### Do not use partial classes
Do not use `partial` classes in the production codebase.

Each type should have a single authoritative implementation file so behavior, invariants, and ownership stay easy to discover and maintain.

### Keep each C# file under 800 lines
To facilitate code porting, each `.cs` file shall contain fewer than 800 lines of code.

When a file approaches this limit, extract cohesive helper types or support classes into separate files instead of continuing to grow a single implementation file.

## Syntax restrictions

### Do not use the `=>` operator
Do not use the `=>` operator in the C# codebase.

Avoid expression-bodied members, lambda expressions, and switch expressions. Use block-bodied members, explicit delegates, loops, and `switch` statements instead so the code stays consistent with this restriction.

## Using directives

### Do not use aliases in `using` directives
Do not use alias directives such as `using Foo = Bar;` in the C# codebase.

Import the namespace directly and use the original type names instead so the code stays explicit and consistent.

## Portability

### Prefer Java-friendly implementation styles
To make future conversion to Java easier, prefer implementation styles that translate cleanly to Java code.

Favor explicit loops, simple control flow, named helper methods, and explicit object construction over C#-specific shorthand or abstractions that do not map directly to Java when a straightforward alternative is practical.
