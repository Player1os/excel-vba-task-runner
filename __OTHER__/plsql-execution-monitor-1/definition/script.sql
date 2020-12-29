-- SQL*PLUS

CREATE TABLE "exec.script" (
	"id" VARCHAR2(256) NOT NULL,
	"execution_interval_sec" NUMBER(14) NOT NULL,
	"execution_offset_sec" NUMBER(14) NOT NULL,
	"last_execution_timestamp" NUMBER(14)
);

ALTER TABLE "exec.script"
	ADD CONSTRAINT "exec.script_pk"
	PRIMARY KEY ("id")
;
