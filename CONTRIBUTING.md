# Contributing to to-xlsx

Thank you for considering contributing to to-xlsx! This document outlines the process for
contributing to this project.

## How Can I Contribute?

### Reporting Bugs

Found a bug? Please create an issue with the following information:

- A clear, descriptive title
- Steps to reproduce the bug
- Expected behavior
- Actual behavior
- Any additional context (screenshots, error messages, etc.)

### Suggesting Enhancements

Have an idea to improve to-xlsx? Please create an issue with:

- A clear, descriptive title
- A detailed description of the enhancement
- Why this enhancement would be useful
- Any implementation ideas you have

### Pull Requests

1. Fork the repository
2. Create a new branch (`git checkout -b feature/your-feature-name`)
3. Make your changes
4. Run tests (`pnpm test`)
5. Commit your changes (`git commit -m 'Add some feature'`)
6. Push to the branch (`git push origin feature/your-feature-name`)
7. Open a Pull Request

## Development Setup

1. Clone the repository
2. Install dependencies with pnpm:
    ```bash
    pnpm install
    ```
3. Make your changes
4. Build the project:
    ```bash
    pnpm build
    ```
5. Test your changes:
    ```bash
    pnpm test
    ```

## Coding Standards

- Use TypeScript for all source code
- Follow the existing code style (we use ESLint and Prettier)
- Write descriptive commit messages
- Update documentation for any API changes

## Testing

- Add tests for new features
- Ensure all tests pass before submitting a PR

## Release Process

The maintainers will handle releases following this process:

1. Update version in package.json
2. Update CHANGELOG.md
3. Create a new release tag
4. Publish to npm

## Questions?

Feel free to open an issue with your question or reach out to the maintainers.

Thank you for contributing to to-xlsx!
