my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            53230A-4
DATE:                  2018-12-20 12:20:23
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       156
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
  1.010  MATH         COLUNA = 5
  1.011  MATH         EX = 0
  1.012  MATH         DIV = 0
#--------------------------------CLEAR---------------------------------
  1.013  ASK-   R D   N B            P J S U       M C X Z        A  L  T  W
  #---------------------- PLANILHA ----------------------
  1.014  DISP         [32]          ATENÇÃO!!!!
  1.014  DISP         [32]
  1.014  DISP         [32]     GPIB CONTADOR 53230A = 3
  1.014  DISP         [32]     GPIB GERADOR 81150A = 1
#---------------------------CONFIG EXCEL--------------------------
  1.015  LIB          COM xlWS = xlApp.Worksheets["PER2"];
  1.016  LIB          xlWS.Select();
#-----------------------------CONFIG  Nº MEAS----------------
  1.017  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.018  MATH         A = RND(MEM)
  1.019  IF           A == 0
  1.020  JMP          1.118
  1.021  ENDIF
#----------------------------CONFIG GENERATOR----------------------
  1.022  RSLT         =
  1.023  IEEE         [@1]*RST
  1.024  IEEE         :ROSC:SOUR EXT
  1.025  IEEE         VOLT:AMPL 1VRMS
  1.026  IEEE         VOLT:OFFS 0
  1.027  IEEE         OUTP1:IMP 50
  1.028  IEEE         OUTP1:LOAD 50
  1.029  IEEE         FUNC SIN
 #---------------------------CONFIG CONT-----------------------
  1.030  IEEE         [@3]*RST
  1.031  IEEE         :FUNC 'PER 2'
  1.032  IEEE         INP2:COUP DC
  1.033  IEEE         INP2:IMP 50
  1.034  IEEE         INP2:FILT ON
  1.035  IEEE         INP2:LEV:AUTO OFF
  1.036  IEEE         INP2:LEV .2
#-----------------SETUP----------------
  1.037  DISP         CONECTE A SAÍDA 1 DO GERADOR A ENTRADA DO CONTADOR 1
  1.037  DISP
  1.037  DISP         [32]   GENERATOR      to         UNIVERSAL COUNTER
  1.037  DISP         [32]
  1.037  DISP         [32]
  1.037  DISP         [32]   OUT 1 -------------------> CHANNEL 2
  1.037  DISP         [32]
  1.037  DISP         [32]
  1.038  PIC          SETUPCH2-1
#-----------------CONFIG POINT------------------
  1.039  DO
  1.040  LIB          COM ED = xlApp.Cells[LP,CP];
  1.041  LIB          PONTO = ED.Value2;
  1.042  IF           PONTO == 0
  1.043  JMP          1.118
  1.044  ENDIF
  1.045  MATH         CP = CP + 1
  1.046  LIB          COM DE = xlApp.Cells[LP,CP];
  1.047  LIB          TEX = DE.Value2;
  1.048  MATH         P = PONTO&TEX
  1.049  MATH         Z1 = CMP  (TEX,"MHz")
  1.050  MATH         Z2 = CMP  (TEX,"kHz")
  1.051  MATH         Z3 = CMP  (TEX,"Hz")
  1.052  IF           Z1 == 1
  1.053  MATH         EX = 1E6
  1.054  ENDIF
  1.055  IF           Z2 == 1
  1.056  MATH         EX = 1E3
  1.057  ENDIF
  1.058  IF           Z3 == 1
  1.059  MATH         EX = 1E0
  1.060  ENDIF
  1.061  IF           Z1 == 1 && PONTO == 1
  1.062  MATH         EX = 1E6
  1.063  ENDIF
  1.064  IF           Z2 == 1 && PONTO == 1
  1.065  MATH         EX = 1E3
  1.066  ENDIF
  1.067  IF           Z3 == 1 && PONTO == 1
  1.068  MATH         EX = 1E0
  1.069  ENDIF
#----------------OUT GEN-------------------
  1.070  IEEE         [@1]FREQ [V P]
  1.071  IEEE         OUTP1 ON
  1.072  WAIT         [D2000]
#----------------------END-------------------------------
  1.073  IF           P == 00
  1.074  JMP          1.118
  1.075  ENDIF
#----------------CONFIG FILTER 100k-------------------
  1.076  IF           PONTO > 100 && Z2 == 1 || Z1 == 1
  1.077  IEEE         [@3]INP2:FILT OFF
  1.078  IEEE         INP2:LEV:AUTO ON
  1.079  ENDIF
#------------DIVISOR----------------
  1.080  IF           Z1 == 1
  1.081  MATH         DIV = 1E-9
  1.082  ENDIF
  1.083  IF           Z2 == 1
  1.084  MATH         DIV = 1E-6
  1.085  ENDIF
  1.086  IF           Z3 == 1
  1.087  MATH         DIV = 1E0
  1.088  ENDIF
  1.089  IF           Z1 == 1 && PONTO == 1
  1.090  MATH         DIV = 1E-6
  1.091  ENDIF
  1.092  IF           Z2 == 1 && PONTO == 1
  1.093  MATH         DIV = 1E-3
  1.094  ENDIF
  1.095  IF           Z3 == 1 && PONTO == 1
  1.096  MATH         DIV = 1E0
  1.097  ENDIF
#------------CONFIG IN METER----------------
  1.098  MATH         TEMPO = GATE + (GATE / 2)
  1.099  IEEE         [@3]FREQ:GATE:TIME [V GATE]
  1.100  IEEE         INIT
  1.101  DO
  1.102  WAIT         -t [V TEMPO] WAIT A MINUTE
  1.103  IEEE         FETCh?[I]
  1.104  MATH         MEM = MEM / DIV
#------------------SAVE DATE----------------
  1.105  LIB          COM selectedCell5 = xlApp.Cells[LINHA,COLUNA];
  1.106  LIB          selectedCell5.Select();
  1.107  LIB          selectedCell5.FormulaR1C1 = [MEM];
  1.108  MATH         T = T + 1
  1.109  MATH         COLUNA = COLUNA + 1
  1.110  MATH         CP = CP + 1
  1.111  UNTIL        T == A
  1.112  MATH         T = 0
  1.113  MATH         COLUNA = 5
  1.114  MATH         LINHA = LINHA + 1
  1.115  MATH         CP = 1
  1.116  MATH         LP = LP + 1
  1.117  UNTIL        P == 0
#------------------RESET------------------
  1.118  IEEE         [@3]*RST
  1.119  IEEE         [@1]*RST
