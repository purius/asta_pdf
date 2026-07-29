namespace PdfMergeTool.Services;

/// <summary>
/// Cancels stale document work and serializes mutations for the currently open PDF.
/// </summary>
public sealed class DocumentOperationCoordinator : IDisposable
{
    private readonly object _sync = new();
    private readonly SemaphoreSlim _mutationGate = new(1, 1);
    private readonly List<CancellationTokenSource> _retiredDocumentCancellations = [];
    private CancellationTokenSource _currentDocumentCancellation = new();
    private int _generation;
    private bool _disposed;

    public int StartNewDocument()
    {
        lock (_sync)
        {
            ThrowIfDisposed();
            _currentDocumentCancellation.Cancel();
            _retiredDocumentCancellations.Add(_currentDocumentCancellation);
            _currentDocumentCancellation = new CancellationTokenSource();
            return ++_generation;
        }
    }

    public DocumentOperationToken Capture()
    {
        lock (_sync)
        {
            ThrowIfDisposed();
            return new DocumentOperationToken(_generation, _currentDocumentCancellation.Token);
        }
    }

    public bool IsCurrent(DocumentOperationToken operation)
    {
        lock (_sync)
        {
            return !_disposed &&
                   operation.Generation == _generation &&
                   !operation.CancellationToken.IsCancellationRequested;
        }
    }

    public void ThrowIfSuperseded(DocumentOperationToken operation)
    {
        operation.CancellationToken.ThrowIfCancellationRequested();
        if (!IsCurrent(operation))
        {
            throw new OperationCanceledException("The document was replaced before this operation completed.");
        }
    }

    public async Task<IDisposable> EnterMutationAsync(DocumentOperationToken operation)
    {
        ThrowIfSuperseded(operation);
        await _mutationGate.WaitAsync(operation.CancellationToken);
        try
        {
            ThrowIfSuperseded(operation);
            return new MutationLease(_mutationGate);
        }
        catch
        {
            _mutationGate.Release();
            throw;
        }
    }

    public void Dispose()
    {
        lock (_sync)
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _currentDocumentCancellation.Cancel();
            _currentDocumentCancellation.Dispose();
            foreach (var cancellation in _retiredDocumentCancellations)
            {
                cancellation.Dispose();
            }

            _retiredDocumentCancellations.Clear();
        }
    }

    private void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }

    private sealed class MutationLease(SemaphoreSlim gate) : IDisposable
    {
        private SemaphoreSlim? _gate = gate;

        public void Dispose()
        {
            Interlocked.Exchange(ref _gate, null)?.Release();
        }
    }
}

public readonly record struct DocumentOperationToken(int Generation, CancellationToken CancellationToken);
