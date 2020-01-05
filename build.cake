#addin "nuget:?package=NuGet.Protocol&version=5.0.2"
#addin "nuget:?package=NuGet.Versioning&version=5.0.2"
#addin "nuget:?package=Cake.ExtendedNuGet"
#addin "nuget:?package=NuGet.Core"
#addin "nuget:?package=Cake.Codecov"
#addin "nuget:?package=Cake.Figlet"

#tool "nuget:?package=xunit.runner.console"
#tool "nuget:?package=JetBrains.dotCover.CommandLineTools"
#tool "nuget:?package=Codecov"

#l "common.cake"

using NuGet;

//////////////////////////////////////////////////////////////////////
// ARGUMENTS
//////////////////////////////////////////////////////////////////////

var projectName = "EPPlus.Core.Extensions";
var solution = "./" + projectName + ".sln";
var testProject = GetFiles($"./test/**/*{projectName}.Tests.csproj").First();

var target = Argument("target", "Default");
var configuration = Argument("configuration", "Release");
var toolpath = Argument("toolpath", @"tools");
var branch = Argument("branch", EnvironmentVariable("APPVEYOR_REPO_BRANCH"));

var nugetApiKey = EnvironmentVariable("nugetApiKey");            

var nupkgPath = "nupkg";
var nupkgRegex = $"**/{projectName}*.nupkg";
var nugetPath = toolpath + "/nuget.exe";
var nugetQueryUrl = "https://www.nuget.org/api/v2/";
var nugetPushUrl = "https://www.nuget.org/api/v2/package";
var NUGET_PUSH_SETTINGS = new NuGetPushSettings
                          {
                              ToolPath = File(nugetPath),
                              Source = nugetPushUrl,
                              ApiKey = nugetApiKey
                          };

//////////////////////////////////////////////////////////////////////
// TASKS
//////////////////////////////////////////////////////////////////////

Setup(context =>
{
	Information(Figlet(projectName));	
});

Task("Clean")
    .Does(() =>
    {
        Information("Current Branch is:" + EnvironmentVariable("APPVEYOR_REPO_BRANCH"));
        CleanDirectories("./src/**/bin");
        CleanDirectories("./src/**/obj");
        CleanDirectory(nupkgPath);
    });

Task("Restore-NuGet-Packages")
    .IsDependentOn("Clean")
    .Does(() =>
    {
        DotNetCoreRestore(solution);
    });

Task("Build")
    .IsDependentOn("Restore-NuGet-Packages")
    .Does(() =>
    {
        DotNetCoreBuild(solution, new DotNetCoreBuildSettings{Configuration = configuration});
    });

Task("Run-Unit-Tests")
    .IsDependentOn("Build")
    .Does(() =>
    {           
		DotCoverAnalyse(tool =>
					tool.DotNetCoreTest(testProject.FullPath, new DotNetCoreTestSettings { Configuration = configuration }), 
				    new FilePath("./coverage.xml"),
				    new DotCoverAnalyseSettings { ReportType = DotCoverReportType.DetailedXML }
					 .WithFilter("+:EPPlus.Core.Extensions")
					 .WithFilter("-:EPPlus.Core.Extensions.Tests")					
					);	   								
    });

Task("Upload-Coverage")
	.IsDependentOn("Run-Unit-Tests")
	.WithCriteria(() => !AppVeyor.Environment.PullRequest.IsPullRequest)
    .Does(() =>
	{
		Codecov(new CodecovSettings {
						Files = new[] { "./coverage.xml" },						
						Token = EnvironmentVariable("COVERALLS_REPO_TOKEN"),
						Branch = branch
			});
	});

Task("Pack")
    .IsDependentOn("Upload-Coverage")
	.WithCriteria(() => branch == "master" && !AppVeyor.Environment.PullRequest.IsPullRequest)
    .Does(() =>
    {
        var nupkgFiles = GetFiles(nupkgRegex);
        MoveFiles(nupkgFiles, nupkgPath);
    });

Task("NugetPublish")
    .IsDependentOn("Pack")
    .WithCriteria(() => branch == "master" && !AppVeyor.Environment.PullRequest.IsPullRequest)
    .Does(()=>
    {
        foreach(var nupkgFile in GetFiles(nupkgRegex))
        {
          if(!IsNuGetPublished(nupkgFile, nugetQueryUrl))
          {
             Information("Publishing... " + nupkgFile);
             NuGetPush(nupkgFile, NUGET_PUSH_SETTINGS);
          }
          else
          {
             Information("Already published, skipping... " + nupkgFile);
          }
        }
    });

//////////////////////////////////////////////////////////////////////
// TASK TARGETS
//////////////////////////////////////////////////////////////////////

Task("Default")
    .IsDependentOn("Build")
    .IsDependentOn("Run-Unit-Tests")
    .IsDependentOn("Pack")
    .IsDependentOn("NugetPublish");
    
//////////////////////////////////////////////////////////////////////
// EXECUTION
//////////////////////////////////////////////////////////////////////

RunTarget(target);