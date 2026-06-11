using PmAi.RdpRemoteLauncher;
using Xunit;

namespace PmAiRdpRemoteLauncher.Tests;

public class UncPathSegmentRepairTests
{
    [Fact]
    public void Repair_fixesCollapsedKonan002Segment()
    {
        var wrong = @"\\192.168.0.101\share\002 加工G\scenarios\a.ardrpa";
        var correct = @"\\192.168.0.101\share\002  加工G\scenarios\a.ardrpa";
        Assert.Equal(correct, UncPathSegmentRepair.Repair(wrong));
    }

    [Fact]
    public void Repair_fixesCollapsedKonan002Segment_inFullUncScenarioPath()
    {
        var wrong =
            @"\\192.168.0.101\共有フォルダ\湖南工場\湖南共有\002 加工G\●配台AIシステム\共有DATA\scenarios\工程別生産計画問い合わせ.ardrpa";
        var correct =
            @"\\192.168.0.101\共有フォルダ\湖南工場\湖南共有\002  加工G\●配台AIシステム\共有DATA\scenarios\工程別生産計画問い合わせ.ardrpa";
        Assert.Equal(correct, UncPathSegmentRepair.Repair(wrong));
    }
}
