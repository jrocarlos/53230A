my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            MAIN
DATE:                  2018-12-20 12:13:18
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       58
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON
#------------------CONFIG PLANILHA---------------
  1.001  DISP         CERTIFIQUE-SE QUE A PLANILHA DE DADOS
  1.001  DISP         FOI DEVIDADEMENTE CRIADA COM OS PONTOS E PADRÕES
  1.001  DISP         UTILIZADOS NESTA CALIBRAÇÃO E ESTÁ SALVA COM
  1.001  DISP         O NOME "53230A" NO ENDEREÇO:
  1.001  DISP         "Z:/Software/PLANILHAS"
#---------------------CONFIG EXCEL-------------------
  1.002  MATH         xlFile = "Z:/Software/PLANILHAS/53230A.xlt"
  1.003  LIB          COM xlApp = "Excel.Application";
  1.004  LIB          xlApp.Visible = True;
  1.005  LIB          COM xlWB = xlApp.Workbooks;
  1.006  LIB          xlWB.Open(xlFile);
#------------------CONFIG WORKSHEET-------------------
  1.007  LIB          COM xlWS = xlApp.Worksheets["CH1"];
  1.008  LIB          xlWS.Select();
#-----------------------CONFIG TEST FREQUENCY CANAL 1------------------------
  1.009  OPBR         DESEJA CALIBRAR FREQUÊNCIA NO CANAL 1?
  1.010  JMPT         1.012
  1.011  JMP          1.013
  1.012  CALL         53230A-1
#-----------------------CONFIG TEST PERIOD CANAL 1------------------------
  1.013  OPBR         DESEJA CALIBRAR PERÍODO NO CANAL 1?
  1.014  JMPT         1.016
  1.015  JMP          1.017
  1.016  CALL         53230A-2
#-----------------------CONFIG TEST FREQUENCY CANAL 2------------------------
  1.017  OPBR         DESEJA CALIBRAR FREQUÊNCIA NO CANAL 2?
  1.018  JMPT         1.020
  1.019  JMP          1.021
  1.020  CALL         53230A-3
#-----------------------CONFIG TEST PERIOD CANAL 2------------------------
  1.021  OPBR         DESEJA CALIBRAR PERÍODO NO CANAL 2?
  1.022  JMPT         1.024
  1.023  JMP          1.025
  1.024  CALL         53230A-4
#-----------------------CONFIG TEST FREQUENCY CANAL 3------------------------
  1.025  OPBR         DESEJA CALIBRAR FREQUÊNCIA NO CANAL 3?
  1.026  JMPT         1.028
  1.027  JMP          1.029
  1.028  CALL         53230A-5
#-----------------------CONFIG TEST REFERENCE------------------------
  1.029  OPBR         DESEJA CALIBRAR FREQUÊNCIA DE REFERÊNCIA?
  1.030  JMPT         1.032
  1.031  JMP          1.033
  1.032  CALL         53230A-6
#-----------------------END CAL------------------------
  1.033  DISP         FIM
