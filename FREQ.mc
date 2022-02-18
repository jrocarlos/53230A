my company                                                  MET/CAL Procedure
=============================================================================
INSTRUMENT:            53230A-1
DATE:                  2018-12-19 16:01:10
AUTHOR:                Carlos Júnior
REVISION:
ADJUSTMENT THRESHOLD:  70%
NUMBER OF TESTS:       1
NUMBER OF LINES:       175
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
  1.014  DISP         [32]     GPIB GERADOR 81150A = 1
  1.014  DISP         [32]     GPIB GERADOR E8257D = 20
#---------------------------CONFIG EXCEL--------------------------
  1.015  LIB          COM xlWS = xlApp.Worksheets["CH1"];
  1.016  LIB          xlWS.Select();
#-----------------------------CONFIG  Nº MEAS----------------
  1.017  MEMI         DIGITE O NÚMERO DE MEDIDAS
  1.018  MATH         A = RND(MEM)
  1.019  IF           A == 0
  1.020  JMP          1.123
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
  1.031  IEEE         :FUNC 'FREQ 1'
  1.032  IEEE         INP1:COUP DC
  1.033  IEEE         INP1:IMP 50
  1.034  IEEE         INP1:FILT ON
  1.035  IEEE         INP:LEV:AUTO OFF
  1.036  IEEE         INP:LEV .2
#-----------------SETUP----------------
  1.037  DISP         CONECTE A SAÍDA 1 DO GERADOR A ENTRADA DO CONTADOR 1
  1.037  DISP
  1.037  DISP         [32]   GENERATOR      to         UNIVERSAL COUNTER
  1.037  DISP         [32]
  1.037  DISP         [32]
  1.037  DISP         [32]   OUT 1 -------------------> CHANNEL 1
  1.037  DISP         [32]
  1.037  DISP         [32]
  1.038  PIC          SETUPCH1-1
#-----------------CONFIG POINT------------------
  1.039  DO
  1.040  LIB          COM P1 = xlApp.Cells[LP,CP];
  1.041  LIB          PONTO = P1.Value2;
  1.042  IF           PONTO == 0
  1.043  JMP          1.123
  1.044  ENDIF
  1.045  MATH         CP = CP + 1
  1.046  LIB          COM T1 = xlApp.Cells[LP,CP];
  1.047  LIB          TEX = T1.Value2;
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
#----------------100 MHz >>-------------------
  1.070  IF           XT == 0 && PONTO > 100 && Z1 == 1
  1.071  JMP          1.117
  1.072  ENDIF
#----------------OUT GEN-------------------
  1.073  IF           XT == 0
  1.074  IEEE         [@1]FREQ [V P]
  1.075  IEEE         OUTP1 ON
  1.076  WAIT         [D2000]
  1.077  ENDIF
#----------------OUT GEN 2-------------------
  1.078  IF           XT == 1
  1.079  IEEE         [@20]FREQ [V P]
  1.080  IEEE         OUTP:STAT ON
  1.081  WAIT         [D2000]
  1.082  ENDIF
#----------------------END-------------------------------
  1.083  IF           P == 00
  1.084  JMP          1.123
  1.085  ENDIF
#----------------CONFIG FILTER 100k-------------------
  1.086  IF           PONTO > 100 && Z2 == 1 || Z1 == 1
  1.087  IEEE         [@3]INP1:FILT OFF
  1.088  IEEE         INP:LEV:AUTO ON
  1.089  ENDIF
#------------CONFIG IN METER----------------
  1.090  MATH         TEMPO = GATE + (GATE / 2)
  1.091  IEEE         [@3]FREQ:GATE:TIME [V GATE]
  1.092  IEEE         INIT
  1.093  DO
  1.094  WAIT         -t [V TEMPO] WAIT A MINUTE
  1.095  IEEE         FETCh?[I]
  1.096  MATH         MEM = MEM / EX
#------------------SAVE DATE----------------
  1.097  LIB          COM selectedCell = xlApp.Cells[LINHA,COLUNA];
  1.098  LIB          selectedCell.Select();
  1.099  LIB          selectedCell.FormulaR1C1 = [MEM];
  1.100  MATH         T = T + 1
  1.101  MATH         COLUNA = COLUNA + 1
  1.102  MATH         CP = CP + 1
  1.103  UNTIL        T == A
  1.104  MATH         T = 0
  1.105  MATH         COLUNA = 3
  1.106  MATH         LINHA = LINHA + 1
  1.107  MATH         CP = 1
  1.108  MATH         LP = LP + 1
  1.109  UNTIL        P == 0
#----------------OUT GEN-------------------
  1.110  IF           XT == 0
  1.111  IEEE         OUTP1 OFF
  1.112  ENDIF
#----------------OUT GEN 2-------------------
  1.113  IF           XT == 1
  1.114  IEEE         OUTP:STAT OFF
  1.115  ENDIF
  1.116  JMP          1.123
#----------------CONFIG >100 MHz----------------------
  1.117  DISP         Connect the generator to the UUT as follows:
  1.117  DISP
  1.117  DISP         [32]   Generator         to         Counter
  1.117  DISP         [32]   RF OUTPUT -------------------> CHANNEL 3
  1.117  DISP         [32]
  1.117  DISP         [32]     GPIB CONTADOR 53230A = 3
  1.117  DISP         [32]     GPIB GERADOR E8257D = 20
  1.118  PIC          SETUPCH1-2
  1.119  IEEE         [@1]*RST
  1.119  IEEE         [@20]*RST
  1.120  IEEE         :ROSC:SOUR:AUTO ON
  1.120  IEEE         :POW:LEV:IMM:AMPL 13
  1.121  MATH         XT = XT + 1
  1.122  JMP          1.073
#------------------RESET------------------
  1.123  IEEE         [@3]*RST
  1.124  IEEE         [@1]*RST
  1.125  IEEE         [@20]*RST
