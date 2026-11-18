using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public sealed record PolarisSourceQuestion(int SourceNumber, string Number, string Question, string Answer);

/// <summary>Source ranges remain Excel serial numbers; children keep their parent's printed number.</summary>
public static class PolarisSubquestionExpander
{
    private sealed record Rule(string Number, int Count, string Hash);
    private static readonly Dictionary<string, IReadOnlyDictionary<int, Rule>> Rules =
        new Dictionary<string, IReadOnlyDictionary<int, Rule>>(StringComparer.Ordinal)
        {
            ["final-polaris-1"] = new Dictionary<int, Rule>
            {
                [284] = new("3-8-C-1", 4, "E59B0A7805EE4F9B88F8F3028312981B0DE51D82BB048C55D59EA4DDFE77A187"),
                [285] = new("3-9", 10, "8771AF69318EF6CFA36F8CEAE1AD2FE9423A9BC8EA16B3B4E6B1BE3107711F14"),
                [286] = new("3-10", 10, "4ABA5C00BA58F6C870ED891BBEF8AC407DB280D565C3B62C2FFB93BC03C6F0CC")
            },
            ["final-polaris-2"] = new Dictionary<int, Rule>(),
            ["final-polaris-3"] = new Dictionary<int, Rule>
            {
                [45] = new("1-5", 5, "D814CBFE1C9DB8F61018ED650AD5DF1539E43FEDA2CB0B142DD6BF23D36C0130"),
                [124] = new("2-3B", 5, "FEE89859E9268DCAEFFE65A20A20824BAFC074FFA6342BE485346D3CFEADA580"),
                [130] = new("2-4B", 5, "8392C0BE3CBF60D7E864FA22C8B577381F8F083DFAD3B62EF66F8AFFDCA42284"),
                [136] = new("2-5B", 5, "B8CC2CF74F614823D01EC1BDFF333BAED3E3D4925A9E956FB7233BB67BB04C4F"),
                [137] = new("2-6A", 5, "29D2AE6E188D64E31ADE1175A1EB6CB7256EB0A2B2BAEF43A05E45FFE004BA9C"),
                [138] = new("2-6B", 5, "27361CA923AC74142E82CE062922E133465CC1B03C7F8E05ACD4DF070276FB50"),
                [272] = new("3-10A", 5, "BD9F53658E848BD5B873FDDF7507F9105EB91CD53105C804E43A6FDAC0F9CE50"),
                [273] = new("3-10B", 5, "41872B9D48F2B52A7DD61906C4A8CA015388203680C86681CD1D3DEFF0513DDB")
            }
        };

    public static bool Supports(string profileId) => Rules.ContainsKey(profileId);

    public static bool Supports(ExcelTemplateProfile profile) => Supports(ExpansionId(profile));
    public static string ExpansionId(ExcelTemplateProfile profile) => profile.QuestionExpansionProfileId.Length > 0
        ? profile.QuestionExpansionProfileId : profile.Id;
    public static int GetAvailableCount(ExcelTemplateProfile profile, int start, int end) => GetAvailableCount(ExpansionId(profile), start, end);

    public static int GetAvailableCount(string profileId, int start, int end)
    {
        var count = end - start + 1;
        if (Rules.TryGetValue(profileId, out var rules))
            count += rules.Where(p => p.Key >= start && p.Key <= end).Sum(p => p.Value.Count - 1);
        return count;
    }

    public static IReadOnlyList<PolarisSourceQuestion> Expand(string profileId, PolarisSourceQuestion source)
    {
        if (!Rules.TryGetValue(profileId, out var rules) || !rules.TryGetValue(source.SourceNumber, out var rule))
            return [source];

        var text = Normalize(source.Question) + "\n" + Normalize(source.Answer);
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
        if (source.Number != rule.Number || hash != rule.Hash)
            throw new InvalidOperationException($"ポラリスの元番号 {source.SourceNumber}（{source.Number}）は調査済みの問題・解答と異なるため、小問分割を停止しました。登録Excelの版を確認してください。続ける場合はこの範囲を外してください。");

        var question = Normalize(source.Question);
        var answer = Normalize(source.Answer);
        var result = new List<PolarisSourceQuestion>();
        if (profileId == "final-polaris-1")
        {
            var first = Regex.Match(question, @"(?m)^\(1\)");
            var statementsText = question[first.Index..];
            var aggregateChoices = statementsText.IndexOf("\n\n①", StringComparison.Ordinal);
            if (aggregateChoices >= 0) statementsText = statementsText[..aggregateChoices];
            var statements = NumberedParts(statementsText);
            var explanations = NumberedParts(answer, lineStartsOnly: true);
            for (var i = 1; i <= rule.Count; i++)
                Add(i, "英文の誤りを訂正しなさい。誤りがなければ「誤りなし」と答えなさい。\n" + statements[i], explanations[i]);
        }
        else if (source.SourceNumber is 272 or 273)
        {
            var body = question[(question.IndexOf("\n\n", StringComparison.Ordinal) + 2)..];
            var parts = NumberedParts(body);
            var answers = NumberedParts(answer);
            var instruction = source.SourceNumber == 272
                ? "文法上不要な語を1語答えなさい。"
                : "不要な語と直後の語を答えなさい。文末の語を除く場合は、不要な語と×を答えなさい。";
            for (var i = 1; i <= rule.Count; i++)
            {
                // A numbered section can include unmarked contextual sentences. Keep them together.
                Add(i, instruction + "\n" + parts[i], answers[i]);
            }
        }
        else if (source.SourceNumber == 45)
        {
            const string choices = "(A) at  (B) by  (C) down  (D) from  (E) in  (F) into  (G) of  (H) on  (I) through  (J) to";
            var answers = Regex.Matches(answer, @"\([A-J]\)\s*[a-z]+").Select(m => m.Value).ToArray();
            if (answers.Length != rule.Count) throw Invalid(source);
            var body = question[(question.LastIndexOf("\n\n", StringComparison.Ordinal) + 2)..];
            for (var i = 1; i <= rule.Count; i++)
            {
                var sentence = ContainingSentence(body, $"({i})");
                // These three blanks share one sentence. Separate its clauses rather than
                // filling the other blanks: another child must never reveal this one's answer.
                sentence = i switch
                {
                    2 => "Many " + body[(body.IndexOf("elderly people are keen", StringComparison.Ordinal))..
                        body.IndexOf(", they have difficulty", StringComparison.Ordinal)] + ".",
                    3 => "They " + body[(body.IndexOf("have difficulty", StringComparison.Ordinal))..
                        body.IndexOf(", which keeps them", StringComparison.Ordinal)] + ".",
                    4 => "This " + body[(body.IndexOf("keeps them", StringComparison.Ordinal))..
                        body.IndexOf("SNS.", body.IndexOf("keeps them", StringComparison.Ordinal), StringComparison.Ordinal)] + "SNS.",
                    _ => sentence
                };
                Add(i, "空所に入る語を選びなさい。\n" + sentence + "\n" + choices, answers[i - 1]);
            }
        }
        else
        {
            var answers = NumberedParts(answer);
            var bodyStart = source.SourceNumber == 138
                ? question.IndexOf('\n') + 1 : question.IndexOf("\n\n", StringComparison.Ordinal) + 2;
            var body = question[bodyStart..];
            var optionsStart = Regex.Match(body, @"\n\s*①a\.");
            if (optionsStart.Success) body = body[..optionsStart.Index];
            for (var i = 1; i <= rule.Count; i++)
            {
                var marker = source.SourceNumber is 124 or 130 ? $"({i})" : ((char)('①' + i - 1)).ToString();
                var sentence = ContainingSentence(body, marker);
                string instruction;
                var childAnswer = answers[i];
                if (source.SourceNumber is 124 or 130)
                    instruction = source.SourceNumber == 124
                        ? "括弧内を並べ替え、2番目と5番目の語の記号を答えなさい。"
                        : "括弧内を並べ替え、2番目と3番目の語の記号を答えなさい。";
                else if (source.SourceNumber == 136 && i is 2 or 5)
                {
                    instruction = "括弧内の語句を並べ替え、記号の順序を答えなさい。";
                    // The source answer already contains the complete ordering; no guessed solution.
                    childAnswer = Regex.Replace(childAnswer, @"^[a-d]", "");
                }
                else
                {
                    instruction = "括弧内を並べ替え、3番目と5番目の語（語句）を答えなさい。";
                    var wordAnswer = Regex.Match(childAnswer, @"^[a-d]\((3rd:.+5th:.+)\)$");
                    if (!wordAnswer.Success) throw Invalid(source);
                    childAnswer = wordAnswer.Groups[1].Value;
                }
                Add(i, instruction + "\n" + sentence, childAnswer);
            }
        }
        if (result.Count != rule.Count) throw Invalid(source);
        return result;

        void Add(int child, string childQuestion, string childAnswer)
        {
            if (string.IsNullOrWhiteSpace(childQuestion) || string.IsNullOrWhiteSpace(childAnswer)) throw Invalid(source);
            result.Add(new(source.SourceNumber, $"{source.Number}({child})", childQuestion.Trim(), childAnswer.Trim()));
        }
    }

    private static string Normalize(string text) => text.Replace("\r\n", "\n", StringComparison.Ordinal).Trim();
    private static InvalidOperationException Invalid(PolarisSourceQuestion source) =>
        new($"ポラリスの元番号 {source.SourceNumber}（{source.Number}）の小問と解答を対応づけられません。登録Excelの版と問題解答リストを確認してください。");

    private static Dictionary<int, string> NumberedParts(string text, bool lineStartsOnly = false)
    {
        var markers = Regex.Matches(text, (lineStartsOnly ? @"(?m)^" : "") + @"\((\d+)\)");
        var result = new Dictionary<int, string>();
        for (var i = 0; i < markers.Count; i++)
        {
            var start = markers[i].Index + markers[i].Length;
            var end = i + 1 < markers.Count ? markers[i + 1].Index : text.Length;
            result[int.Parse(markers[i].Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture)] = text[start..end].Trim();
        }
        return result;
    }

    private static string ContainingSentence(string text, string marker)
    {
        var position = text.IndexOf(marker, StringComparison.Ordinal);
        if (position < 0 || text.IndexOf(marker, position + marker.Length, StringComparison.Ordinal) >= 0)
            throw new InvalidOperationException($"小問 {marker} の位置を一意に判定できません。登録Excelの版を確認してください。");
        var start = 0;
        var depth = 0;
        for (var index = 0; index < text.Length; index++)
        {
            if (text[index] == '(') depth++;
            if (text[index] == ')') depth--;
            if (depth != 0 || text[index] is not ('.' or '?' or '!')) continue;
            if (text[index] == '.' && Regex.IsMatch(text[..(index + 1)], @"(?:\b[A-Z]|\b(?:Mr|Mrs|Ms|Dr|Prof|St))\.$")) continue;
            var end = index + 1;
            while (end < text.Length && text[end] is '"' or '”') end++;
            if (end < text.Length && !char.IsWhiteSpace(text[end])) continue;
            if (position < end) return text[start..end].Trim();
            start = end;
        }
        return text[start..].Trim();
    }
}
