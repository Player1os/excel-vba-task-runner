-- SQL*PLUS

INSERT INTO "exec.script" (
	"id",
	"execution_interval_sec",
	"execution_offset_sec"
) VALUES (
	'&1', -- SCRIPT_ID
	&2, -- EXECUTION_INTERVAL_SEC
	&3 -- EXECUTION_OFFSET_SEC
);

COMMIT;
