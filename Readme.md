
    DupeLocatr
    ----------
    A menu-driven duplicate file finder & cleaner for Windows.

    - Recursively scans the current folder (and every subfolder) for
      duplicate files of ANY type: images, videos, music, text, documents,
      archives, whatever.
    - Uses a 3-stage "fast path" so it never wastes time hashing files that
      can't possibly be duplicates:
          1) group by file size            (instant, in-memory)
          2) quick partial hash (64 KB)     (cheap, filters out non-matches)
          3) full hash (MD5/SHA256)         (only run on real candidates)
      This means a folder full of large videos gets scanned fast, because
      most videos never even get fully read.
    - NEVER deletes or moves anything without asking first.
    - Deletions go to the Recycle Bin (recoverable), not permanent removal.
    - "Move to folder" uses robocopy under the hood (fast, reliable, built
      into Windows, handles long paths and retries better than a plain copy).
    - Colored console output + a progress bar during scanning.

    Usage:
        powershell -ExecutionPolicy Bypass -File DupeLocatr.ps1
        powershell -ExecutionPolicy Bypass -File DupeLocatr.ps1 -Path "D:\Photos"

    (Or just double-click Run_DupeLocatr.bat sitting next to this file.)
