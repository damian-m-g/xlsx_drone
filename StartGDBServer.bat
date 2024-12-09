Rem compile unit_tests.exe
cmd /c rake c
Rem run the gdbserver
start gdbserver localhost:3333 temp/unit_tests.exe