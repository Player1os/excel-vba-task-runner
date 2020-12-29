-- SQL*PLUS

CREATE TABLE "exec.log" (
	"script_id" VARCHAR2(256) NOT NULL,
	"status" CHAR(1) NOT NULL,
	"timestamp" NUMBER(14) NOT NULL
);

ALTER TABLE "exec.log"
	ADD CONSTRAINT "exec.log_fk1"
	FOREIGN KEY ("script_id")
	REFERENCES "exec.script" ("id")
;
