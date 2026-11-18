using AutomaticTestPrinting.Core.Models;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace AutomaticTestPrinting.Core.Services;

public static class PdfMaterialCatalog
{
    public static PdfRequestPreparation Prepare(NormalTestRequestCandidate request,
        IReadOnlyCollection<string> folders, IReadOnlyList<PdfMaterialProfile>? profiles = null,
        IReadOnlyList<string>? knownPdfCandidates = null)
    {
        try
        {
            var usesSharedCatalog = profiles is null;
            profiles ??= PdfMaterialProfileStore.Load();
            var name = PdfMaterialProfileStore.NormalizeName(request.MaterialName);
            var matches = profiles.Where(profile => profile.Aliases.Append(profile.DisplayName)
                .Any(alias => PdfMaterialProfileStore.NormalizeName(alias) == name)).ToArray();
            if (matches.Length == 0)
            {
                var grouped = usesSharedCatalog ? SharedMaterialCatalog.Load().Items.FirstOrDefault(item =>
                    item.AutoEnabled && item.PdfVariants.Count > 0 && item.Aliases.Append(item.DisplayName)
                        .Any(alias=>PdfMaterialProfileStore.NormalizeName(alias)==name)) : null;
                if (grouped is not null)
                    return new(true,false,"この教材はSTEP・部などを区別して指定してください。登録名：" +
                        string.Join(" ／ ",grouped.PdfVariants.Select(p=>p.DisplayName)));
                var pending = usesSharedCatalog ? SharedMaterialCatalog.Load().Items.FirstOrDefault(item => item.Kind != "Excel" && !item.AutoEnabled &&
                    item.Aliases.Append(item.DisplayName).Any(alias => PdfMaterialProfileStore.NormalizeName(alias) == name)) : null;
                if (pending is not null)
                {
                    if (ExcelMaterialNameMatcher.FindBestMatches(ExcelTemplateProfileStore.LoadDefault().Profiles, request.MaterialName).Count > 0)
                        return new(false, false, "登録済みのExcel教材です。");
                    return new(true, false, string.IsNullOrWhiteSpace(pending.RegistrationReviewReason)
                        ? "PDF教材はカタログ登録済みですが、問題番号とページの対応などの作成設定が未完了です。「PDF教材を登録」で対応表・使用ファイルを確認して登録してください。"
                        : "PDF教材の登録は保留中です：" + pending.RegistrationReviewReason + " 資料を確認後、「PDF教材を登録」で設定してください。");
                }
                if (name.StartsWith("段階突破", StringComparison.Ordinal))
                    return new(true, false, $"この段階突破PDFは未登録です：{request.MaterialName}。科目・年度・段階・回が一致するPDFを用意し、「PDF教材を登録」で段階突破として登録してください。別年度・別の回は自動選択しません。");
                return new(false, false, "未登録のPDF教材です。");
            }
            if (matches.Length != 1) return new(true, false, "PDF教材名が複数の登録に一致しています。");
            var bounds = request.Range.Split('-');
            if (bounds.Length != 2 || !int.TryParse(bounds[0], out var start) || !int.TryParse(bounds[1], out var end) || start < 1 || end < start)
                return new(true, false, "範囲を「開始番号-終了番号」の形で確認してください。");
            var profile = matches[0];
            if (profile.IsStageTest && (start != 1 || end != 1 || request.QuestionCount != 1))
                return new(true, false, "段階突破は範囲1-1・問題数1で全ページを作成します。");
            var inRange = profile.Questions.Where(question => question.Number >= start && question.Number <= end).OrderBy(question => question.Number).ToArray();
            if (profile.AllowSparseRanges && end > profile.Questions.Max(question => question.Number))
                return new(true, false, $"このPDFの練習問題の項目番号は{profile.Questions.Max(question => question.Number)}までです。原本の範囲を確認してください。");
            if (!profile.AllowSparseRanges && (long)end - start + 1 != inRange.Length)
                return new(true, false, $"指定範囲 {start}-{end} に未登録の問題番号があります。登録済み：{string.Join("、", profile.Questions.Select(q => q.Number).Distinct().Order().Take(30))}。原本の範囲を確認し、必要なら「PDF教材を登録」で対応表を補完してください。");
            if (inRange.Length == 0)
                return new(true, false, "指定範囲には、このPDFに収録された練習問題がありません。原本と登録済みの項目番号を確認してください。");
            if (request.QuestionCount < 1 || request.QuestionCount > inRange.Length)
                return new(true, false, $"問題数は1-{inRange.Length}問で指定してください。");
            var files = knownPdfCandidates ?? folders.Where(Directory.Exists).Distinct(StringComparer.OrdinalIgnoreCase)
                .SelectMany(folder => Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories))
                .Where(path => path.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                .Select(Path.GetFullPath).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            var resolvedPaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string Resolve(string fileName)
            {
                if (resolvedPaths.TryGetValue(fileName, out var path)) return path;
                var candidates = knownPdfCandidates is null ? files : files.Where(candidate =>
                    string.Equals(Path.GetFileName(candidate), fileName, StringComparison.OrdinalIgnoreCase) &&
                    File.Exists(candidate)).ToArray();
                path = PdfFileResolutionService.Resolve(fileName, candidates);
                resolvedPaths.Add(fileName, path);
                return path;
            }
            var resolved = inRange.Select(question => new ResolvedPdfQuestion(question.Number, Resolve(question.FileName), question.Pages,
                string.IsNullOrWhiteSpace(question.AnswerFileName) ? null : Resolve(question.AnswerFileName), question.AnswerPages, question.ExpectedPageCount)).ToArray();
            foreach (var group in resolved.Where(q => q.ExpectedPageCount.HasValue).GroupBy(q => q.Path))
            {
                var actual = PdfMaterialGenerationService.GetPageCount(group.Key);
                if (group.Any(q => q.ExpectedPageCount != actual))
                    return new(true, false, $"PDF総ページ数が対応表と違います：{Path.GetFileName(group.Key)}（実際：{actual}ページ）。版・ファイルを確認してください。");
                foreach (var question in group) PdfMaterialGenerationService.ParsePages(question.Pages, actual);
            }
            return new(true, true, profile.IsStageTest
                ? "段階突破PDF連携準備OK：全ページを作成します（問題・解答を分けません）"
                : $"PDF連携準備OK：範囲 {start}-{end}／{request.QuestionCount}問（範囲から抽選）" +
                    (profile.AllowSparseRanges ? "／PDFに未収録の項目番号は除外します" : ""), resolved);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or InvalidOperationException or ArgumentException or System.Text.Json.JsonException or PdfReaderException)
        {
            return new(true, false, $"PDF教材設定を確認してください：{exception.Message}");
        }
    }

    public static ExcelPdfGenerationResult Generate(TestGenerationPlan plan, string outputFolder)
    {
        var available = plan.PdfQuestions ?? throw new ArgumentException("PDFの対応表がありません。", nameof(plan));
        if (plan.QuestionCount < 1 || plan.QuestionCount > available.Count)
            throw new ArgumentException("問題数が登録範囲外です。", nameof(plan));
        var shuffled = available.ToArray();
        Random.Shared.Shuffle(shuffled);
        var selected = shuffled.Take(plan.QuestionCount).OrderBy(question => question.Number).ToArray();
        Directory.CreateDirectory(outputFolder);
        var paths = TestPdfOutputPaths.FindAvailable(outputFolder, plan.StudentName, plan.MaterialName,
            plan.Range, plan.QuestionCount, selected.All(question => question.AnswerPath is not null));
        var created = new List<string>(2);

        void Export(string destination, bool isAnswer)
        {
            using var output = new PdfDocument();
            var exportedPages = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
            var inputs = new Dictionary<string, PdfDocument>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (var question in selected)
                {
                    var source = isAnswer ? question.AnswerPath! : question.Path;
                    if (!exportedPages.TryGetValue(source, out var seenPages))
                    {
                        seenPages = [];
                        exportedPages.Add(source, seenPages);
                    }
                    // Dense question maps can refer to the same source hundreds of times.
                    if (!inputs.TryGetValue(source, out var input))
                    {
                        input = PdfReader.Open(source, PdfDocumentOpenMode.Import);
                        inputs.Add(source, input);
                    }
                    if (!isAnswer && question.ExpectedPageCount.HasValue && question.ExpectedPageCount != input.PageCount)
                        throw new InvalidOperationException($"PDF総ページ数が対応表と違います：{source}。版・ファイルを確認してください。");
                    try
                    {
                        var pages = PdfMaterialGenerationService.ParsePages(isAnswer ? question.AnswerPages : question.Pages, input.PageCount);
                        foreach (var page in pages)
                        {
                            if (seenPages.Add(page))
                                output.AddPage(input.Pages[page - 1]);
                        }
                    }
                    catch (ArgumentException exception)
                    {
                        throw new InvalidOperationException($"{plan.MaterialName}／問題番号 {question.Number}／PDF：{source}：{exception.Message}", exception);
                    }
                }
                using var stream = new FileStream(destination, FileMode.CreateNew, FileAccess.Write);
                created.Add(destination);
                output.Save(stream, closeStream: false);
            }
            finally
            {
                foreach (var input in inputs.Values) input.Dispose();
            }
        }
        try
        {
            Export(paths.ProblemPdfPath, false);
            if (paths.AnswerPdfPath is not null) Export(paths.AnswerPdfPath, true);
            return paths;
        }
        catch
        {
            foreach (var path in created)
            {
                try { File.Delete(path); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
            throw;
        }
    }
}
