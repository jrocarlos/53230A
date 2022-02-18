my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            53230A-5
DATE:                  2018-12-20 12:07:19
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       118
=============================================================================
 STEP    FSC    RANGE NOMINAL        TOLERANCE     MOD1        MOD2  3  4 CON
 #--------------------------------VARIÁVEIS---------------------------------
  1.001  MATH         A = 0
  1.002  MATH         GATE = 10
  1.003  MATH         P = 0
  1.004  MATH         LP = 2
  1.005  MATH         CP = 1
  1.006  MATH         T  = 0
  1.007  MATH         L = 0
  1.008  MATH         LINHA = 2
  1.009  MATH         TEMPO = 0
  1.010  MATH         COLUNA = 3
  1.011  MATH         EX = 0
  1.012  MATH         XT = 0
#--------------------------------CLEAR---------------------------------
  1.013  ASK-   R D   N B            P J S U       M C X Z        A  L  T  W
  #---------------------- PLANILHA ----------------------
  1.014  DISP         [32]          ATENÇÃO!!!!
  1.014  DISP         [32]
  1.014  DISP         [32]     GPIB CONTADOR 53230A = 3
  1.014  DISP         [32]     GPIB GERADOR E8257D = 20
#---------------------------CONFIG EXCEL--------------------------
  1.015  LIB          COM xlWS = xlApp.Worksheets["CH3"];
  1.016  LIB          xlWS.Select();
#-----------------------------CONFIG  Nº MEAS----------------
  1.017  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.018  MATH         A = RND(MEM)
  1.019  IF           A == 0
  1.020  JMP          1.081
  1.021  ENDIF
#----------------------------CONFIG GENERATOR----------------------
  1.022  RSLT         =
  1.023  IEEE         [@20]*RST
  1.024  IEEE         :ROSC:SOUR:AUTO ON
  1.025  IEEE         :POW:LEV:IMM:AMPL 13
 #---------------------------CONFIG CONT-----------------------
  1.026  IEEE         [@3]*RST
  1.027  IEEE         :FUNC 'FREQ 3'
#-----------------SETUP----------------
  1.028  DISP         CONECTE A SAÍDA DO GERADOR A ENTRADA 3 DO CONTADOR
  1.028  DISP
  1.028  DISP         [32]   GENERATOR      to         UNIVERSAL COUNTER
  1.028  DISP         [32]
  1.028  DISP         [32]
  1.028  DISP         [32]   RF OUT  -------------------> CHANNEL 3
  1.028  DISP         [32]
  1.028  DISP         [32]
  1.029  PIC          SETUPCH3
#-----------------CONFIG POINT------------------
  1.030  DO
  1.031  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.032  LIB          PONTO = P1.Value2;
  1.033  IF           PONTO == 0
  1.034  JMP          1.081
  1.035  ENDIF
  1.036  MATH         CP = CP + 1
  1.037  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.038  LIB          TEX = T1.Value2;
  1.039  MATH         P = PONTO&TEX
  1.040  MATH         Z1 = CMP  (TEX,"MHz")
  1.041  MATH         Z2 = CMP  (TEX,"GHz")
  1.042  IF           Z1 == 1
  1.043  MATH         EX = 1E6
  1.044  ENDIF
  1.045  IF           Z2 == 1
  1.046  MATH         EX = 1E9
  1.047  ENDIF
  1.048  IF           Z1 == 1 && PONTO == 1
  1.049  MATH         EX = 1E6
  1.050  ENDIF
  1.051  IF           Z2 == 1 && PONTO == 1
  1.052  MATH         EX = 1E9
  1.053  ENDIF
#----------------OUT GEN-------------------
  1.054  IEEE         [@20]FREQ [V P]
  1.055  IEEE         OUTP:STAT ON
  1.056  WAIT         [D2000]
#----------------------END-------------------------------
  1.057  IF           P == 00
  1.058  JMP          1.081
  1.059  ENDIF
#------------CONFIG IN METER----------------
  1.060  MATH         TEMPO = GATE + (GATE / 2)
  1.061  IEEE         [@3]FREQ:GATE:TIME [V GATE]
  1.062  IEEE         INIT
  1.063  DO
  1.064  WAIT         -t [V TEMPO] WAIT A MINUTE
  1.065  IEEE         FETCh?[I]
  1.066  MATH         MEM = MEM / EX
#------------------SAVE DATE----------------
  1.067  LIB          COM selectedCell4 = xlApp.Cells[LINHA,COLUNA];
  1.068  LIB          selectedCell4.Select();
  1.069  LIB          selectedCell4.FormulaR1C1 = [MEM];
  1.070  MATH         T = T + 1
  1.071  MATH         COLUNA = COLUNA + 1
  1.072  MATH         CP = CP + 1
  1.073  UNTIL        T == A
  1.074  MATH         T = 0
  1.075  MATH         COLUNA = 3
  1.076  MATH         LINHA = LINHA + 1
  1.077  MATH         CP = 1
  1.078  MATH         LP = LP + 1
  1.079  UNTIL        P == 0
#----------------OUT GEN 2-------------------
  1.080  IEEE         [@20]OUTP:STAT OFF
#------------------RESET------------------
  1.081  IEEE         [@3]*RST
  1.082  IEEE         [@20]*RST
