Rem comiple unit_tests.exe
cmd /c rake c
Rem run the gdbserver
start gdbserver localhost:2159 test/unit_tests.exe