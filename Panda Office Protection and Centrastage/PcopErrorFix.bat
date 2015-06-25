@echo off

:stopnprv
sc stop nnsprv
goto stoppicc

:stoppicc
sc stop nnspicc
goto stoptlsc

:stoptlsc
sc stop nnstlsc
goto stopalpc

:stopalpc
sc stop nnsalpc
goto stopstrm

:stopstrm
sc stop nnsstrm
goto stopnprt

:stopnprt
sc stop nnsprot
goto stopids

:stopids
sc stop nnsids
goto startallfw

:startallfw
sc start nnsprv
sc start nnspicc
sc start nnstlsc
sc start nnsalpc
sc start nnsstrm
sc start nnsprot
sc start nnsids
goto stopprot

:stopprot
sc stop psinprot
goto stopfile

:stopfile
sc stop psinfile
goto stopaflt

:stopaflt
sc stop psinaflt
goto stopproc

:stopproc
sc stop psinproc
goto startall2

:startall2
sc start psinproc
sc start psinaflt
sc start psinfile
sc start psinprot
goto END

:END