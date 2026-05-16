# Course Website (MkDocs Material)

## Local setup

```bash
pip install -r website/requirements.txt
mkdocs serve -f website/mkdocs.yml
```

## Build

```bash
mkdocs build -f website/mkdocs.yml
```

## Versioning (mike)

```bash
mike deploy --config-file website/mkdocs.yml --push 2026-spring latest
mike set-default --config-file website/mkdocs.yml --push latest
```
