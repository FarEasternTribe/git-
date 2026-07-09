param(
  [Parameter(Position = 0)]
  [ValidateSet(
    'rules',
    'task',
    'journal',
    'journal-local',
    'journal-todos',
    'commands',
    'orchestrator',
    'google-sync',
    'google-sync-dry',
    'google-auth',
    'paper-search',
    'experiment-note',
    'experiment-note-latest',
    'experiment-onenote-day',
    'sync-command-log',
    'log-command',
    'migration-check',
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
$BundledPython = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$VenvPython = Join-Path $Workspace '.venv\Scripts\python.exe'

function Get-AgentPython {
  if (Test-Path -LiteralPath $BundledPython) {
    return $BundledPython
  }
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
      exit $LASTEXITCODE
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

function Get-AgentDeviceLabel {
  # agent_onenote_logger.py の detect_device_label と同じ判定。
  # Lenovo側で実行しても命令したLogがDesktop扱いにならないようにする。
  $configured = $env:AGENT_DEVICE_LABEL
  if (-not [string]::IsNullOrWhiteSpace($configured)) {
    return $configured.Trim('[', ']')
  }
  $computerName = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { '' }
  $upperName = $computerName.ToUpperInvariant()
  if ($upperName -like '*LENOVO*') { return 'Lenovo' }
  if ($upperName -like '*DESKTOP*') { return 'Desktop' }
  if ($computerName) { return $computerName }
  return 'UnknownPC'
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
    & powershell -NoProfile -ExecutionPolicy Bypass -File '.\append_onenote_command_log.ps1' `
      -Device (Get-AgentDeviceLabel) `
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
    Invoke-OrchestratorTask -Request ("日誌を実行して rawtext を要約する force " + ($Rest -join ' '))
    $latestJournal = Get-LatestJournalFile
    Write-AutoCommandLog `
      -Summary '日誌を実行' `
      -Actions '司令塔Agent経由で日誌Agentを実行; rawtextから日誌Markdownを生成; 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録' `
      -Files "summarize_note5.py; $latestJournal" `
      -Verification '日誌実行コマンドが終了コード0で完了' `
      -NextSteps '必要に応じて実験PPT・Google Tasks・OneNoteへ反映'
  }
  'journal-local' {
    Invoke-OrchestratorTask -Request ("外部APIなし ローカル日誌を実行して rawtext を要約する force " + ($Rest -join ' '))
    $latestJournal = Get-LatestJournalFile
    Write-AutoCommandLog `
      -Summary '外部APIなしでローカル日誌を実行' `
      -Actions '司令塔Agent経由で日誌Agentをローカル要約モード実行; rawtextから日誌Markdownを生成; Google Tasks手動投入用TodoListを生成; 末尾に生データを貼付; 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録' `
      -Files "summarize_note5.py; $latestJournal; 日誌/GoogleTasks手動投入用.md" `
      -Verification 'ローカル日誌実行コマンドが終了コード0で完了' `
      -NextSteps 'Google Tasksへは自動送信せず、手動投入用TodoListからユーザーが確認して貼り付ける'
  }
  'journal-todos' {
    Invoke-OrchestratorTask -Request ("日誌を実行して rawtext を要約し Google Tasks も同期する force " + ($Rest -join ' '))
    $latestJournal = Get-LatestJournalFile
    Write-AutoCommandLog `
      -Summary '日誌を実行しGoogle Tasks同期も実行' `
      -Actions '司令塔Agent経由で日誌Agentを実行; 日誌生成とTODO同期を処理; 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録' `
      -Files "summarize_note5.py; $latestJournal; 日誌/google_todo_synced.jsonl; 日誌/google_todo_queue.jsonl" `
      -Verification '日誌+Google Tasks同期コマンドが終了コード0で完了'
  }
  'commands' {
    Invoke-OrchestratorTask -Request ("rawtext内の@コマンドを実行して Google Tasks も同期する " + ($Rest -join ' '))
    Write-AutoCommandLog `
      -Summary 'rawtext内の@コマンドを実行' `
      -Actions '司令塔Agent経由で日誌Agentの@コマンド実行を起動; Google Tasks同期; 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録' `
      -Files 'summarize_note5.py; 日誌/command_execution_log.md; 日誌/google_todo_synced.jsonl; 日誌/google_todo_queue.jsonl' `
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
      -Files 'setup_google_tasks.py; credentials.json; token_google_tasks.json' `
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
    Invoke-OrchestratorTask -Request ("実験PPT作成 OneNote 2026実験 実験 " + ($Rest -join ' '))
    Write-AutoCommandLog `
      -Summary 'OneNote実験ノートをExperiment.pptxへ転記' `
      -Actions "司令塔Agent経由で実験ノートAgentを実行; 引数: $($Rest -join ' '); 検証AgentとAgent会議を実施; 実行結果を命令したLogへ自動記録" `
      -Files 'Experiment.pptx; append_onenote_experiment_day_to_ppt.ps1; agent_workspace/実験ノートAgent/onenote_to_ppt' `
      -Verification 'OneNote実験ノートからのPPT転記が終了コード0で完了'
  }
  'sync-command-log' {
    Push-Location $Workspace
    try {
      & powershell -NoProfile -ExecutionPolicy Bypass -File '.\sync_onenote_command_log.ps1' @Rest
      if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
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
        exit $LASTEXITCODE
      }
    } finally {
      Pop-Location
    }
  }
  'migration-check' {
    Invoke-AgentPython -CommandArgs (@('.\migration_check.py') + $Rest)
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






