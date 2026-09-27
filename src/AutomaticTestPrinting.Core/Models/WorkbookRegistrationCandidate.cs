namespace AutomaticTestPrinting.Core.Models;

public sealed record WorkbookRegistrationCandidate(
    string WorkbookPath,
    string WorkbookFileName,
    string SuggestedDisplayName,
    IReadOnlyList<string> MatchingProfileIds,
    int SuggestedMaximumQuestionNumber,
    bool CanRegister,
    bool RequiresReview,
    string Message);

public sealed record WorkbookDiscoveryResult(
    IReadOnlyList<WorkbookRegistrationCandidate> Candidates,
    int RegisteredWorkbookCount,
    int UnreadableWorkbookCount);
