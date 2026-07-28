# PDF Merge Tool Context

PDF Merge Tool handles local PDF merge and page-editing workflows on Windows.

## Language

**Split Output Folder**:
The folder created next to a source PDF to hold files produced by a split operation. It never overwrites an existing folder; repeated operations use a numbered sibling folder such as `document_split (2)`.
_Avoid_: split folder, output directory

**Split Batch**:
A consecutive group of up to N source pages written as one file by an every-N-pages split. The final batch is written even when it contains fewer than N pages.
_Avoid_: page range, split unit

**Parity Split**:
A split operation that writes the odd-numbered source pages and even-numbered source pages into separate PDFs. If a parity has no pages, its output PDF is omitted and the completion notice reports zero files for that parity.
_Avoid_: odd-even split, alternating split

**Multi-file Split**:
Applying one split command to several selected source PDFs. Each source file is processed independently; a failure is recorded for that source while later files continue to process.
_Avoid_: bulk split, batch split

**Split Interval**:
The positive integer N used by every-N-pages split. It defaults to 1, accepts values larger than the source page count, and rejects non-integers and values below 1.
_Avoid_: page count, split size
