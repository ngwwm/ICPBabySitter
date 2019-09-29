# ICPBabySitter
g++ -fPIC -I"$JAVA_HOME/include" -I"$JAVA_HOME/include/linux" -shared -o ./bin/libha4crypto.so ha4crypto.cpp

g++ -std=gnu++11 ha4passwd.cpp -o ./bin/ha4passwd

g++ -std=gnu++11 -D_LICENSE_ ha4passwd.cpp -o ./bin/ha4pwdlic
