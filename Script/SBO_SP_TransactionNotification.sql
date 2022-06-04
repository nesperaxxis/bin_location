-- B1 DEPENDS: BEFORE:PT:PROCESS_START

ALTER PROCEDURE SBO_SP_TransactionNotification
(
	in object_type nvarchar(30), 				-- SBO Object Type
	in transaction_type nchar(1),			-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
	in num_of_cols_in_key int,
	in list_of_key_cols_tab_del nvarchar(255),
	in list_of_cols_val_tab_del nvarchar(255)
)
LANGUAGE SQLSCRIPT
AS
-- Return values
error  int;				-- Result (0 for no error)
error_message nvarchar (200); 		-- Error string to be displayed
begin

error := 0;
error_message := N'Ok';

--------------------------------------------------------------------------------------------------------------------------------

--	ADD	YOUR	CODE	HERE
IF(:object_type = '15' AND :transaction_type = 'A') THEN

	UPDATE "T1"
	SET "T1"."U_AXC_BINLocation" = "T2"."BINCode"
	FROM "ODLN" "T0"
	LEFT JOIN "DLN1" "T1" ON "T1"."DocEntry" = "T0"."DocEntry"
	LEFT JOIN "AXXIS_SAPB1"."AXXIS_TB_BINAutoSelect" "T2" ON "T2"."DocNum" = "T0"."DocNum" 
	AND "T2"."Series" = "T0"."Series" AND (CAST("T2"."RowNum" AS BIGINT) - 1) = "T1"."LineNum"
	WHERE "T0"."DocEntry" = :list_of_cols_val_tab_del
	AND "T2"."CompanyDB" = current_schema;

END IF;

--------------------------------------------------------------------------------------------------------------------------------

-- Select the return values
select :error, :error_message FROM dummy;

end;

