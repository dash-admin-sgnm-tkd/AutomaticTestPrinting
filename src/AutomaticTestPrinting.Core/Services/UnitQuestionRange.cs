using System.Globalization;
using AutomaticTestPrinting.Core.Models;

namespace AutomaticTestPrinting.Core.Services;

public static class UnitQuestionRange
{
    public static (int Start, int End) Resolve(object[,] mapping, int startUnit, int endUnit)
    {
        if (startUnit < 1 || endUnit < startUnit || endUnit > mapping.GetLength(0))
            throw new InvalidOperationException("単元番号が対応表の範囲外です。開始・終了の単元を確認してください。");
        var first = 0;
        var last = 0;
        for (var unit = startUnit; unit <= endUnit; unit++)
        {
            var row = mapping.GetLowerBound(0) + unit - 1;
            var column = mapping.GetLowerBound(1);
            if (!int.TryParse(Convert.ToString(mapping[row,column],CultureInfo.InvariantCulture),out var start) ||
                !int.TryParse(Convert.ToString(mapping[row,column+1],CultureInfo.InvariantCulture),out var end) ||
                start < 1 || end < start || (unit > startUnit && start != last + 1))
                throw new InvalidOperationException($"単元 {unit} の問題番号対応を確認できません。元Excelの対応リストを確認してください。");
            if (unit == startUnit) first = start;
            last = end;
        }
        return (first,last);
    }

    public static int CandidateCount(ExcelTemplateProfile profile,int start,int end)
    {
        if (profile.RangeUsesSectionMapping || profile.QuestionListUsesSectionNumbers)
            return profile.RangeQuestionCounts.Count>0
                ? profile.RangeQuestionCounts.Where(pair=>pair.Key>=start && pair.Key<=end).Sum(pair=>pair.Value)
                : profile.EffectiveMaximumQuestionCount;
        return PolarisSubquestionExpander.GetAvailableCount(profile,start,end) -
            profile.UnavailableQuestionNumbers.Count(number=>number>=start && number<=end) +
            profile.AdditionalQuestionCounts.Where(pair=>pair.Key>=start && pair.Key<=end).Sum(pair=>pair.Value);
    }
}
