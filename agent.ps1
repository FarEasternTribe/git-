param(
  [Parameter(Position = 0)]
  [ValidateSet(
    'rules',
    'task',
    'journal',
    'journal-local',
    'journal-todos',
    'journal-task-status',
    'commands',
    'orchestrator',
    'google-sync',
    'google-sync-dry',
    'google-auth',
    'paper-search',
    'paper-index',
    'experiment-note',
    'experiment-note-latest',
    'experiment-onenote-day',
    'experiment-board',
    'sync-command-log',
    'log-command',
    'migration-check',
    'status',
    'secrets-status',
    'backup',
    'restore-check',
    'test',
    'daily-command-log-check',
    'mutual-check',
    'device-monitor'
  )]
  [string]$Action = 'rules',

  [Parameter(ValueFromRemainingArguments = $true)]
  [string[]]$Rest
)

$ErrorActionPreference = 'Stop'

$Workspace = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $Workspace '.venv\Scripts\python.exe'

function Get-AgentPython {
  if (Test-Path -LiteralPath $VenvPython) {
    return $VenvPython
  }
  return 'python'
}

function Invoke-AgentPython {
  param(
    [Parameter(Mandatory = $true)]
    [AllowEmptyString()]
    [string[]]$CommandArgs
  )

  $CommandArgs = @($CommandArgs | Where-Object { -not [string]::IsNullOrEmpty($_) })
  $Python = Get-AgentPython
  Push-Location $Workspace
  try {
    & $Python @CommandArgs
    if ($LASTEXITCODE -ne 0) {
      throw "Agent command failed with exit code $LASTEXITCODE"
    }
  } finally {
    Pop-Location
  }
}

function Invoke-OrchestratorTask {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Request,
    [string[]]$ExtraArgs = @()
  )

  $orchestratorArgs = @('.\orchestrator_agent.py', '--execute', '--run', $Request)
  if ($null -ne $ExtraArgs) {
    $orchestratorArgs += @($ExtraArgs | Where-Object { -not [string]::IsNullOrEmpty($_) })
  }
  $orchestratorArgs = @($orchestratorArgs | Where-Object { -not [string]::IsNullOrEmpty($_) })
  Invoke-AgentPython -CommandArgs $orchestratorArgs
}

function Get-LatestJournalFile {
  $journalDir = Join-Path $Workspace '日誌'
  if (-not (Test-Path -LiteralPath $journalDir)) {
    return ''
  }
  $latest = Get-ChildItem -LiteralPath $journalDir -Filter '*_日誌*.md' -File |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
  if ($null -eq $latest) {
    return ''
  }
  return $latest.FullName
}

function Write-AutoCommandLog {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Summary,
    [string]$Actions = '',
    [string]$Files = '',
    [string]$Verification = '実行成功',
    [string]$RequiredOnOtherDevice = '相互監視Agentが必要性を判断する',
    [string]$NextSteps = '必要に応じて次のAgent処理へつなぐ'
  )

  Push-Location $Workspace
  try {
    # -Device は指定しない: append_onenote_command_log.ps1 が COMPUTERNAME から
    # Desktop/Lenovo を自動判定する（旧Codex版はここで 'Desktop' 固定だったため
    # Lenovo機でも誤ってDesktopと記録されていた）。
    & powershell -NoProfile -ExecutionPolicy Bypass -File '.\append_onenote_command_log.ps1' `
      -Summary $Summary `
      -Actions $Actions `
      -Files $Files `
      -Verification $Verification `
      -RequiredOnOtherDevice $RequiredOnOtherDevice `
      -NextSteps $NextSteps
    if ($LASTEXITCODE -ne 0) {
      Write-Warning "Auto command log failed with exit code $LASTEXITCODE"
    }
  } catch {
    Write-Warning "Auto command log failed: $($_.Exception.Message)"
  } finally {
    Pop-Location
  }
}

function Start-AgentRun {
  param([string]$RunAction, [string[]]$RunArguments)
  if ($RunAction -in @('status', 'secrets-status', 'restore-check')) {
    return [pscustomobject]@{ Skip = $false; RunId = ''; Detail = 'read-only utility' }
  }
  $Python = Get-AgentPython
  $dedupeActions = @('journal', 'journal-local', 'journal-todos', 'commands', 'google-sync', 'paper-search', 'paper-index', 'experiment-note', 'experiment-note-latest', 'experiment-onenote-day', 'experiment-board')
  $dedupeSeconds = if ($RunAction -in $dedupeActions) { 600 } else { 0 }
  $force = @($RunArguments | Where-Object { $_ -match '^(--force|-Force|強制)$' }).Count -gt 0
  $runtimeArgs = @('.\agent_runtime.py', 'begin', '--action', $RunAction, '--dedupe-seconds', [string]$dedupeSeconds)
  $argumentText = ($RunArguments -join ' ').Trim()
  if (-not [string]::IsNullOrWhiteSpace($argumentText)) { $runtimeArgs += @("--arguments=$argumentText") }
  if ($force) { $runtimeArgs += '--force' }
  Push-Location $Workspace
  try {
    $resultText = (& $Python @runtimeArgs | Out-String).Trim()
    $code = $LASTEXITCODE
  } finally {
    Pop-Location
  }
  if ($code -eq 10) {
    return [pscustomobject]@{ Skip = $true; RunId = ''; Detail = $resultText }
  }
  if ($code -ne 0) { throw "Could not start runtime record (exit $code)" }
  $result = $resultText | ConvertFrom-Json
  return [pscustomobject]@{ Skip = $false; RunId = $result.run_id; Detail = $resultText }
}

function Complete-AgentRun {
  param([string]$RunId, [string]$Status, [int]$ExitCode, [string]$Message = '')
  if ([string]::IsNullOrWhiteSpace($RunId)) { return }
  $Python = Get-AgentPython
  Push-Location $Workspace
  try {
    & $Python '.\agent_runtime.py' finish --run-id $RunId --status $Status --exit-code $ExitCode --message $Message *> $null
  } finally {
    Pop-Location
  }
}

$Run = Start-AgentRun -RunAction $Action -RunArguments $Rest
if ($Run.Skip) {
  Write-Host "同じ操作が直近10分以内に正常完了しているため、重複実行を停止しました。強制実行が必要な場合は --force を明示してください。"
  return
}

try {
switch ($Action) {
  'rules' {
    Invoke-AgentPython -CommandArgs @('.\summarize_note5.py', '--show-rules')
  }
  'task' {
    $request = ($Rest -join ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($request)) {
      throw 'task action requires a natural-language request.'
    }
    Invoke-OrchestratorTask -Request $request
    Write-AutoCommandLog `
      -Summary '司令塔Agent経由でユーザー依頼を実行' `
      -Actions "orchestrator_agent.py --execute --run を実行; 依頼: $request; 担当Agent選定・実行・検証・Agent会議を実施" `
      -Files 'orchestrator_agent.py; verification_agent.py; tools/orchestrator_agent_log.jsonl; agent_workspace' `
      -Verification '司令塔Agent経由の処理が終了コード0で完了'
  }
  'journal' {
    Invoke-OrchestratorTask -Request ("外部APIなし ローカル日誌を差分実行して rawtext を要約する " + ($Rest -join ' '))
    $latestJournal = Get-LatestJournalFile
    Write-AutoCommandLog `
      -Summary '日誌を実行' `
      -Actions '司令塔Agent経由で日誌Agentを実行; rawtextから日誌Markdownを生成; 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録' `
      -Files "summarize_note5.py; $latestJournal" `
      -Verification '日誌実行コマンドが終了コード0で完了' `
      -NextSteps '必要に応じて実験PPT・Google Tasks・OneNoteへ反映'
  }
  'journal-local' {
    Invoke-OrchestratorTask -Request ("外部APIなし ローカル日誌を差分実行して rawtext を要約する " + ($Rest -join ' '))
    $latestJournal = Get-LatestJournalFile
    Write-AutoCommandLog `
      -Summary '外部APIなしでローカル日誌を実行' `
      -Actions '司令塔Agent経由で日誌Agentをローカル要約モード実行; rawtextから日誌Markdownを生成; Google Tasks手動投入用TodoListを生成; 末尾に生データを貼付; 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録' `
      -Files "summarize_note5.py; $latestJournal; 日誌/GoogleTasks手動投入用.md" `
      -Verification 'ローカル日誌実行コマンドが終了コード0で完了' `
      -NextSteps 'Google Tasksへは自動送信せず、手動投入用TodoListからユーザーが確認して貼り付ける'
  }
  'journal-todos' {
    # ユーザー方針: Google TasksへのAPI送信は停止。日誌はローカル要約し、
    # @todoはコピペ用の手動投入用リストにまとめる（journalと同じ動作）。
    Invoke-OrchestratorTask -Request ("外部APIなし ローカル日誌を差分実行して rawtext を要約する " + ($Rest -join ' '))
    $latestJournal = Get-LatestJournalFile
    Write-AutoCommandLog `
      -Summary '日誌を実行（@todoは手動投入用リストに集約）' `
      -Actions '司令塔Agent経由で日誌Agentをローカル要約モード実行; 日誌生成; @todoをGoogle Tasks手動投入用リストに集約（API送信なし）; 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録' `
      -Files "summarize_note5.py; $latestJournal; 日誌/GoogleTasks手動投入用.md" `
      -Verification '日誌実行コマンドが終了コード0で完了' `
      -NextSteps 'Google Tasksへは自動送信せず、手動投入用リストからユーザーが確認して貼り付ける'
  }
  'journal-task-status' {
    Push-Location $Workspace
    try {
      & powershell -NoProfile -ExecutionPolicy Bypass -File '.\get_onenote_journal_tasks.ps1' @Rest
      if ($LASTEXITCODE -ne 0) {
        throw "Journal task-status command failed with exit code $LASTEXITCODE"
      }
    } finally {
      Pop-Location
    }
  }
  'commands' {
    Invoke-OrchestratorTask -Request ("rawtext内の@コマンドを実行する " + ($Rest -join ' '))
    Write-AutoCommandLog `
      -Summary 'rawtext内の@コマンドを実行' `
      -Actions '司令塔Agent経由で日誌Agentの@コマンド実行を起動; @todoは手動投入用リストに集約（Google Tasks送信なし）; @askは無効; 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録' `
      -Files 'summarize_note5.py; 日誌/command_execution_log.md; 日誌/GoogleTasks手動投入用.md' `
      -Verification '@コマンド実行が終了コード0で完了'
  }
  'orchestrator' {
    Invoke-AgentPython -CommandArgs (@('.\orchestrator_agent.py') + $Rest)
    Write-AutoCommandLog `
      -Summary '司令塔Agentを実行' `
      -Actions "orchestrator_agent.py を実行; 依頼引数: $($Rest -join ' '); 実行結果を命令したLogへ自動記録" `
      -Files 'orchestrator_agent.py; tools/orchestrator_agent_log.jsonl; agent_workspace/司令塔Agent' `
      -Verification '司令塔Agentコマンドが終了コード0で完了'
  }
  'google-sync' {
    Invoke-OrchestratorTask -Request ("Google Tasksキュー同期を実行 " + ($Rest -join ' '))
    Write-AutoCommandLog `
      -Summary 'Google Tasksキュー同期を実行' `
      -Actions 'sync_google_todo_queue.py を実行; キュー内TODOをGoogle Tasksへ同期; 実行結果を命令したLogへ自動記録' `
      -Files 'sync_google_todo_queue.py; 日誌/google_todo_queue.jsonl; 日誌/google_todo_synced.jsonl' `
      -Verification 'Google Tasks同期コマンドが終了コード0で完了'
  }
  'google-sync-dry' {
    Invoke-OrchestratorTask -Request ("Google Tasksキュー同期をドライランで確認だけ実行 " + ($Rest -join ' '))
  }
  'google-auth' {
    Invoke-AgentPython -CommandArgs (@('.\setup_google_tasks.py') + $Rest)
    Write-AutoCommandLog `
      -Summary 'Google Tasks認証確認を実行' `
      -Actions 'setup_google_tasks.py を実行; Google Tasks OAuth設定を確認; 実行結果を命令したLogへ自動記録' `
      -Files 'setup_google_tasks.py; %LOCALAPPDATA%\OpenAI-Agent\secrets\credentials.json; %LOCALAPPDATA%\OpenAI-Agent\secrets\token_google_tasks.json' `
      -Verification 'Google Tasks認証コマンドが終了コード0で完了'
  }
  'paper-search' {
    Invoke-OrchestratorTask -Request ("論文検索を実行 " + ($Rest -join ' '))
    Write-AutoCommandLog `
      -Summary '論文検索を実行' `
      -Actions '司令塔Agent経由で論文検索Agentを実行; 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録' `
      -Files 'orchestrator_agent.py; tools/paper_search_agent.ps1; papers' `
      -Verification '論文検索コマンドが終了コード0で完了'
  }
  'paper-index' {
    Invoke-OrchestratorTask -Request ("paper-index 論文PDFライブラリの要約とINDEX化を実行 " + ($Rest -join ' '))
    Write-AutoCommandLog `
      -Summary '論文PDFライブラリ索引を更新' `
      -Actions "司令塔Agent経由でPaperIndexAgent(ローカル・API不使用)を実行; PDF集約・Markdown索引・MASTER_INDEX/Excel/CSV更新; 検証AgentとAgent会議を実施; 引数: $($Rest -join ' ')" `
      -Files 'orchestrator_agent.py; paper_index_agent.py; verification_agent.py; papers/library/PDFs; papers/library/index/MASTER_INDEX.md; papers/library/summaries' `
      -Verification 'paper-indexコマンドが終了コード0で完了; MASTER_INDEX.md生成を確認'
  }
  'experiment-note' {
    Invoke-OrchestratorTask -Request ("実験ノートをExperiment.pptxへ追記 " + ($Rest -join ' '))
    Write-AutoCommandLog `
      -Summary '実験ノートをExperiment.pptxへ追記' `
      -Actions "司令塔Agent経由で実験ノートAgentを実行; 引数: $($Rest -join ' '); 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録" `
      -Files 'Experiment.pptx; orchestrator_agent.py; append_onenote_experiment_day_to_ppt.ps1; agent_workspace/実験ノートAgent' `
      -Verification '実験PPT追記コマンドが終了コード0で完了'
  }
  'experiment-note-latest' {
    Invoke-OrchestratorTask -Request ("実験ノートをExperiment.pptxへ追記 latest " + ($Rest -join ' '))
    Write-AutoCommandLog `
      -Summary '最新日誌から実験ノートをExperiment.pptxへ追記' `
      -Actions '司令塔Agent経由で実験ノートAgentを実行; 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録' `
      -Files 'Experiment.pptx; orchestrator_agent.py; agent_workspace/実験ノートAgent' `
      -Verification '最新日誌からの実験PPT追記が終了コード0で完了'
  }
  'experiment-onenote-day' {
    Invoke-OrchestratorTask -Request ("実験PPT作成 OneNote FarEasternTribe 実験 " + ($Rest -join ' '))
    Write-AutoCommandLog `
      -Summary 'OneNote実験ノートをExperiment.pptxへ転記' `
      -Actions "司令塔Agent経由で実験ノートAgentを実行; 引数: $($Rest -join ' '); 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録" `
      -Files 'Experiment.pptx; append_onenote_experiment_day_to_ppt.ps1; agent_workspace/実験ノートAgent/onenote_to_ppt' `
      -Verification 'OneNote実験ノートからのPPT転記が終了コード0で完了'
  }
  'experiment-board' {
    Invoke-OrchestratorTask -Request ("進行中の実験ボードを生成 " + ($Rest -join ' '))
    Write-AutoCommandLog `
      -Summary '進行中実験ボードを生成' `
      -Actions "司令塔Agent経由で実験ノートAgent(トラッカー)を実行; 引数: $($Rest -join ' '); 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録" `
      -Files 'experiment_tracker.ps1; agent_workspace/実験ノートAgent/experiment_board/latest_experiment_board.md' `
      -Verification '進行中実験ボードの生成が終了コード0で完了'
  }
  'sync-command-log' {
    Push-Location $Workspace
    try {
      & powershell -NoProfile -ExecutionPolicy Bypass -File '.\sync_onenote_command_log.ps1' @Rest
      if ($LASTEXITCODE -ne 0) {
        throw "Command-log sync failed with exit code $LASTEXITCODE"
      }
    } finally {
      Pop-Location
    }
  }
  'log-command' {
    Push-Location $Workspace
    try {
      & powershell -NoProfile -ExecutionPolicy Bypass -File '.\append_onenote_command_log.ps1' @Rest
      if ($LASTEXITCODE -ne 0) {
        throw "Command logging failed with exit code $LASTEXITCODE"
      }
    } finally {
      Pop-Location
    }
  }
  'migration-check' {
    Invoke-AgentPython -CommandArgs (@('.\migration_check.py') + $Rest)
  }
  'status' {
    Invoke-AgentPython -CommandArgs (@('.\agent_runtime.py', 'status') + $Rest)
  }
  'secrets-status' {
    Invoke-AgentPython -CommandArgs @('.\agent_runtime.py', 'secrets-status')
  }
  'backup' {
    Invoke-AgentPython -CommandArgs (@('.\agent_runtime.py', 'backup') + $Rest)
  }
  'restore-check' {
    if ($Rest.Count -lt 1) { throw 'restore-check requires a backup ZIP path.' }
    Invoke-AgentPython -CommandArgs (@('.\agent_runtime.py', 'restore') + $Rest)
  }
  'test' {
    Invoke-AgentPython -CommandArgs @('-m', 'unittest', 'discover', '-s', 'tests', '-v')
  }
  'daily-command-log-check' {
    Invoke-OrchestratorTask -Request ("DesktopとLenovoの相互監視、命令したLog差分確認、再現性チェックを実行 " + ($Rest -join ' '))
  }
  'mutual-check' {
    Invoke-OrchestratorTask -Request ("device-monitor mutual-check: Desktop and Lenovo rule-change check. Sync OneNote OpenAI_Agent1 命令したLog, inspect changed command-log pages, identify Lenovo/Desktop rule or workflow changes, apply necessary safe local code/runbook/settings updates on this device, rerun verification and Agent council when fixes are made, summarize changed items, and record the final result to 命令したLog. " + ($Rest -join ' '))
  }
  'device-monitor' {
    Invoke-OrchestratorTask -Request ("DesktopとLenovoの相互監視、命令したLog差分確認、再現性チェックを実行 " + ($Rest -join ' '))
  }
}
  Complete-AgentRun -RunId $Run.RunId -Status 'succeeded' -ExitCode 0 -Message 'completed'
} catch {
  Complete-AgentRun -RunId $Run.RunId -Status 'failed' -ExitCode 1 -Message $_.Exception.Message
  throw
}






