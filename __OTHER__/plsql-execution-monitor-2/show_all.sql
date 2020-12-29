SELECT
	t.*,
	t.rowid
FROM
	"exec.script" t
;

SELECT
	t.*,
	t.rowid
FROM
	"exec.log" t
ORDER BY
	t."timestamp" DESC
;
