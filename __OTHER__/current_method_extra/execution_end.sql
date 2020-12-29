INSERT INTO
	"exec.log"
VALUES (
	'&1', -- SCRIPT_ID
	'E',
	TO_CHAR(SYSDATE, 'YYYYMMDDHH24MISS')
)
;
UPDATE
	"exec.script"
SET
	"last_execution_timestamp" = TO_CHAR(SYSDATE, 'YYYYMMDDHH24MISS')
WHERE
	"id" = '&1' -- SCRIPT_ID
;
COMMIT
