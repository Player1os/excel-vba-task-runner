# Excel VBA Task Runner

[![License](https://img.shields.io/github/license/Player1os/excel-vba-task-runner.svg)](https://github.com/Player1os/excel-vba-task-runner/blob/master/LICENSE)

Contains the `execute.vbs` script for executing external task scripts while offering the following features:

- Checks whether the task script is available (i.e. exists and is reachable).
- Creates log files that store the task script's standard output and standard error steams during each task run.
- Creates an additional log file which contains additional information about the task run.
- Checks whether the returned status code of the task script is equal to zero.
- Maintains an upper limit on the amount of stored task run logs.

If the task script is not available, the script waits a specified amount of seconds before trying again. This repeats for a specified
amount of times. If all attempts fail, the error is reported to the user via an email message.

If the returned status code of the task script is not equal to zero, it is reported to the user via an email message.

If an email message cannot be successfully sent for any reason, the user is notified via a message box.

## Log file storage

All log files are stored within a specified path. The file system hierarchy created within that path is as follows:

1. The name of the task that is being executed.
2. The timestamp of the start of the task's execution.

Under the final directory, three files are created during the task run:

- `out.log` which contains the standard output of the task script.
- `err.log` which contains the standard error of the task script.
- `info.log` which contains details about the current task run:
	- The name of the user who triggered the task run.
	- The name of the machine where the task was executed.
	- The path to the task script file.
	- The timestamp fo the start of the task run.
	- The timestamp fo the end of the task run.
	- The status code returned by the task script.

After each task run, the current task log directory is check, to see if the specified maximum task log count has been exceeded. If true,
the oldest extraneous task run log directories are removed.

## Usage instructions

The `execute.vbs` script takes two parameters:

1. The name of the task to be used for identification.
2. The path to the task script to be executed.

### Example

```batch
execute.vbs task-name C:\path\to\task\script.bat
```

## Manually task killing

To manually kill a running task, execute the following command:

```batch
kill.bat task-name
```

## Deployment instructions

1. Create a copy of the `.env.reset.bat` script, name it `.env.set.bat` and modify it as follows:
	- Set `DEPLOY_DIRECTORY_PATH` to a suitable location, where the project's scripts will be deployed.
	- Set `APP_TASK_RUNNER_WAIT_SECONDS` to the number of seconds between each attempt at checking the availability of the
	task script file.
	- Set `APP_TASK_RUNNER_ITERATION_COUNT` to the number of attempts at checking the availability of the task script file.
	- Set `APP_TASK_RUNNER_LOG_DIRECTORY_PATH` to a suitable location, where all log files will be stored as detailed above.
	- Set `APP_TASK_RUNNER_MAXIMUM_TASK_LOG_COUNT` to the maximum number of task run logs to be stored per each task.
	- Set `APP_TASK_RUNNER_ERROR_MAIL_RECIPIENT` to the email address (or semicolon separated list of email addresses), where an email
	message will be sent once an error occurs.
2. Run the `deploy.bat` script to copy the required files to the deploy directory.
