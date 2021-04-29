import cx_Oracle
import os

os.environ['PATH'] = 'C:\Program Files (x86)\Oracle\instantclient_19_10'
dsn_tns = cx_Oracle.makedsn('fetldb1', '1524', service_name='PCTTLIV')
conn = cx_Oracle.connect(user='fpc', password='fpc', dsn=dsn_tns)
print('Connect to FPC Oracle DB')


cur = conn.cursor()
query = """
        SELECT U.FLOW_ID
      , U.PROCESS
      , U.PROC_DISP
      , U.PRODUCT_MODEL
      , U.SIDE
      , U.REJ_ID
      , U.REJ_CODE
      , U.REJ_NAME
      , U.PERIOD
      , U.START_DATE
      , U.END_DATE
      , U.LOT_COUNT
      , U.INPUT_QTY
      , U.REJECT_QTY
      , U.SUM_P
      , U.N_BAR
      , U.P_BAR
      , CASE WHEN U.UCL = 0 OR U.UCL < U.STD_UCL THEN
                  U.STD_UCL
             ELSE
                  U.UCL
        END AS UCL
      , CASE WHEN U.UCL = 0 OR U.UCL < U.STD_UCL THEN
                  U.STD_CL
             ELSE
                  U.CL
        END AS CL
      , CASE WHEN U.UCL = 0 OR U.UCL < U.STD_UCL THEN
                  U.STD_LCL
             ELSE
                  U.LCL
        END AS LCL
FROM (SELECT UL.FLOW_ID AS FLOW_ID
             , UL.PROC_ID AS PROCESS
             , PR.PROC_DISP
             , UL.PRD_MODEL AS PRODUCT_MODEL
             , UL.SIDE AS SIDE
             , UL.REJ_ID AS REJ_ID
             , RJ.REJ_CODE AS REJ_CODE
             , RJ.REJ_DESC AS REJ_NAME
             , UL.PERIOD
             , UL.START_DATE
             , UL.END_DATE
             , UL.LOT_COUNT AS LOT_COUNT
             , UL.INPUT_QTY
             , UL.REJECT_QTY
             , UL.SUM_P
             , UL.N_BAR
             , UL.P_BAR
             , ROUND((UL.P_BAR + (3*ROUND(SQRT(UL.P_BAR*(1 - UL.P_BAR))/ROUND(SQRT(UL.N_BAR),10),10)))*100,5) AS UCL
             , ROUND(UL.P_BAR*100,5) AS CL
             , CASE WHEN ROUND((UL.P_BAR - (3*ROUND(SQRT(UL.P_BAR*(1 - UL.P_BAR))/ROUND(SQRT(UL.N_BAR),10),10)))*100,5) > 0 THEN
                         ROUND((UL.P_BAR - (3*ROUND(SQRT(UL.P_BAR*(1 - UL.P_BAR))/ROUND(SQRT(UL.N_BAR),10),10)))*100,5)
                    ELSE 0
               END AS LCL
             , ROUND((UL.P_BAR_STD + (3*ROUND(SQRT(UL.P_BAR_STD*(1 - UL.P_BAR_STD))/ROUND(SQRT(UL.N_BAR),10),10)))*100,5) AS STD_UCL
             , ROUND(UL.P_BAR_STD*100,5) AS STD_CL
             , CASE WHEN ROUND((UL.P_BAR_STD - (3*ROUND(SQRT(UL.P_BAR_STD*(1 - UL.P_BAR_STD))/ROUND(SQRT(UL.N_BAR),10),10)))*100,5) > 0 THEN
                         ROUND((UL.P_BAR_STD - (3*ROUND(SQRT(UL.P_BAR_STD*(1 - UL.P_BAR_STD))/ROUND(SQRT(UL.N_BAR),10),10)))*100,5)
                    ELSE 0
               END AS STD_LCL
      FROM
      (SELECT '0010' AS FLOW_ID
             , M.PERIOD
             , M.START_DATE
             , M.END_DATE
             , SUBSTR(A.AOI_PRD_NAME,1,INSTR(A.AOI_PRD_NAME,'-'))||SUBSTR(SUBSTR(A.AOI_PRD_NAME,INSTR(A.AOI_PRD_NAME,'-')+1,10),1,INSTR(SUBSTR(A.AOI_PRD_NAME,INSTR(A.AOI_PRD_NAME,'-')+1,10),'-') - 1) AS PRD_MODEL
             , A.AOI_PROC_ID AS PROC_ID
             , A.AOI_SIDE AS SIDE
             , A.AOI_REJ_CODE AS REJ_ID
             , COUNT(*) AS LOT_COUNT
             , SUM(A.AOI_PCS) AS INPUT_QTY
             , SUM(A.AOI_REJ_QTY) AS REJECT_QTY
             , SUM(ROUND((A.AOI_REJ_QTY/A.AOI_PCS),10)) AS SUM_P
             , ROUND(AVG(A.AOI_PCS),0) AS N_BAR
             , ROUND(AVG(A.AOI_REJ_QTY)/ROUND(AVG(A.AOI_PCS),0),10) AS P_BAR
             , ROUND(2/ROUND(AVG(A.AOI_PCS),0),10) AS P_BAR_STD
      FROM FPC_AOI_INSPECTION A
           ,( SELECT '0010' AS FLOW_ID
                     , MM.PERIOD
                     , MM.START_DATE
                     , MM.END_DATE
                     , MM.PREV_DATE
                     , SUBSTR(AA.AOI_PRD_NAME,1,INSTR(AA.AOI_PRD_NAME,'-'))||SUBSTR(SUBSTR(AA.AOI_PRD_NAME,INSTR(AA.AOI_PRD_NAME,'-')+1,10),1,INSTR(SUBSTR(AA.AOI_PRD_NAME,INSTR(AA.AOI_PRD_NAME,'-')+1,10),'-') - 1) AS PRD_MODEL
                     , AA.AOI_PROC_ID AS PROC_ID
                     , AA.AOI_SIDE AS SIDE
                     , AA.AOI_REJ_CODE AS REJ_ID
                     , COUNT(*) AS LOT_COUNT
              FROM FPC_AOI_INSPECTION AA
                   ,(
                      SELECT DISTINCT  TO_CHAR(TO_DATE('01/01/2021','DD/MM/YYYY') + (LEVEL - 1),'MM/YYYY') AS PERIOD
                                       , TO_DATE('01/'||TO_CHAR(TO_DATE('01/01/2021','DD/MM/YYYY') + (LEVEL - 1),'MM/YYYY'),'DD/MM/YYYY') AS START_DATE
                                       , TO_DATE('25/'||TO_CHAR(TO_DATE('01/01/2021','DD/MM/YYYY') + (LEVEL - 1),'MM/YYYY'),'DD/MM/YYYY') AS END_DATE
                                       , ADD_MONTHS(TO_DATE('01/'||TO_CHAR(TO_DATE('01/01/2021','DD/MM/YYYY') + (LEVEL - 1),'MM/YYYY'),'DD/MM/YYYY'),-1) AS PREV_DATE
                      FROM DUAL
                      CONNECT BY LEVEL <= ((TO_DATE('28/02/2021','DD/MM/YYYY') - TO_DATE('01/01/2021','DD/MM/YYYY')) + 1)
                    ) MM
      WHERE AA.AOI_PRD_NAME LIKE 'RG%'
            AND AA.AOI_DATE >= MM.START_DATE
            AND AA.AOI_DATE <=  MM.END_DATE
            AND AA.AOI_PROC_ID IS NOT NULL
      GROUP BY SUBSTR(AA.AOI_PRD_NAME,1,INSTR(AA.AOI_PRD_NAME,'-'))||SUBSTR(SUBSTR(AA.AOI_PRD_NAME,INSTR(AA.AOI_PRD_NAME,'-')+1,10),1,INSTR(SUBSTR(AA.AOI_PRD_NAME,INSTR(AA.AOI_PRD_NAME,'-')+1,10),'-') - 1)
               , AA.AOI_PROC_ID
               , AA.AOI_SIDE
               , AA.AOI_REJ_CODE
               , MM.PERIOD
               , MM.START_DATE
               , MM.END_DATE
               , MM.PREV_DATE
             ) M
      WHERE SUBSTR(A.AOI_PRD_NAME,1,INSTR(A.AOI_PRD_NAME,'-'))||SUBSTR(SUBSTR(A.AOI_PRD_NAME,INSTR(A.AOI_PRD_NAME,'-')+1,10),1,INSTR(SUBSTR(A.AOI_PRD_NAME,INSTR(A.AOI_PRD_NAME,'-')+1,10),'-') - 1) = M.PRD_MODEL
            AND A.AOI_SIDE = M.SIDE
            AND A.AOI_PROC_ID = M.PROC_ID
            AND A.AOI_REJ_CODE = M.REJ_ID
            AND ((M.LOT_COUNT >= 32
            AND A.AOI_DATE >= M.START_DATE
            AND A.AOI_DATE <= M.END_DATE)
            OR  (M.LOT_COUNT < 32
            AND A.AOI_DATE >= M.PREV_DATE
            AND A.AOI_DATE <= M.END_DATE ))
      GROUP BY SUBSTR(A.AOI_PRD_NAME,1,INSTR(A.AOI_PRD_NAME,'-'))||SUBSTR(SUBSTR(A.AOI_PRD_NAME,INSTR(A.AOI_PRD_NAME,'-')+1,10),1,INSTR(SUBSTR(A.AOI_PRD_NAME,INSTR(A.AOI_PRD_NAME,'-')+1,10),'-') - 1)
               , A.AOI_PROC_ID
               , A.AOI_SIDE
               , A.AOI_REJ_CODE
               , M.PERIOD
               , M.START_DATE
               , M.END_DATE
      HAVING COUNT(*) >= 32
       ) UL
      , FPC_REJECT_MASTER RJ
      , FPC_PROCESS PR
 WHERE UL.REJ_ID = RJ.REJ_ID
       AND UL.PROC_ID = PR.PROC_ID
       AND UL.LOT_COUNT >= 32
 ORDER BY UL.PRD_MODEL ASC
          , UL.PROC_ID ASC
          , UL.SIDE ASC
          , UL.REJ_ID ASC
) U
"""
cur.execute(query)

row_no = 0
for row in cur:
    print(row)
    
    row_no += 1
    if row_no > 100:
        break

conn.close()
print('Disconnect to FPC Oracle DB')