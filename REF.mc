my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            53230A-6
DATE:                  2020-01-09 16:05:21
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       96
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
  1.014  DISP         [32]     GPIB CONTADOR 53132A = 5
#---------------------------CONFIG EXCEL--------------------------
  1.015  LIB          COM xlWS = xlApp.Worksheets["REF"];
  1.016  LIB          xlWS.Select();
#-----------------------------CONFIG  Nº MEAS----------------
  1.017  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.018  MATH         A = RND(MEM)
  1.019  IF           A == 0
  1.020  JMP          1.063
  1.021  ENDIF
 #---------------------------CONFIG CONT-----------------------
  1.022  IEEE         [@5]*RST
  1.023  IEEE         :FUNC 'FREQ 1'
  1.024  IEEE         INIT:CONT OFF
  1.025  IEEE         INP1:COUP DC
  1.026  IEEE         INP1:IMP 50
  #1.027  IEEE         INP1:FILT ON
  1.027  IEEE         EVEN1:LEV 0.5V
#-----------------SETUP----------------
  1.028  DISP         CONECTE A SAÍDA 1 DO GERADOR A ENTRADA DO CONTADOR 1
  1.028  DISP
  1.028  DISP         [32]   GENERATOR      to         UNIVERSAL COUNTER
  1.028  DISP         [32]
  1.028  DISP         [32]
  1.028  DISP         [32]   OUT 1 -------------------> CHANNEL 1
  1.028  DISP         [32]
  1.028  DISP         [32]
  1.029  PIC          SETUPREF
#-----------------CONFIG POINT------------------
  1.030  DO
  1.031  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.032  LIB          PONTO = P1.Value2;
  1.033  IF           PONTO == 0
  1.034  JMP          1.063
  1.035  ENDIF
  1.036  MATH         CP = CP + 1
  1.037  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.038  LIB          TEX = T1.Value2;
  1.039  MATH         P = PONTO&TEX
#----------------------END-------------------------------
  1.040  IF           P == 00
  1.041  JMP          1.063
  1.042  ENDIF
#------------CONFIG IN METER----------------
  1.043  MATH         TEMPO = GATE + (GATE / 2)
  1.044  IEEE         [@5]FREQ:ARM:STOP:TIM [V GATE]
  1.045  IEEE         INIT:CONT ON
  1.046  DO
  1.047  WAIT         -t [V TEMPO] WAIT A MINUTE
  1.048  IEEE         FETCh?[I]
  1.049  MATH         MEM = MEM / 1E6
#------------------SAVE DATE----------------
  1.050  LIB          COM selectedCell7 = xlApp.Cells[LINHA,COLUNA];
  1.051  LIB          selectedCell7.Select();
  1.052  LIB          selectedCell7.FormulaR1C1 = [MEM];
  1.053  MATH         T = T + 1
  1.054  MATH         COLUNA = COLUNA + 1
  1.055  MATH         CP = CP + 1
  1.056  UNTIL        T == A
  1.057  MATH         T = 0
  1.058  MATH         COLUNA = 3
  1.059  MATH         LINHA = LINHA + 1
  1.060  MATH         CP = 1
  1.061  MATH         LP = LP + 1
  1.062  UNTIL        P == 0
#------------------RESET------------------
  1.063  IEEE         [@5]*RST
