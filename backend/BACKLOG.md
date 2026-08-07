# Backend Backlog

Things we've deliberately deferred, to come back to later.

- **Port the 50MB size cap + page-by-page PDF compression rescue to IL petty cash** (`/extract-petty-cash`). Currently only applied to GA petty cash and GA ProdCC — IL still rejects any file over 40MB outright. Same fix (`_check_and_compress_pdf_size` / `_compress_oversized_pdf_pages` in `main.py`) should drop in the same way once IL is ready.
