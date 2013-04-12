#!/bin/bash

echo "world del all" | yarp rpc /icubSim/world
#echo "world mk sph 0.1     -0.15 1.0 0.5      1 0 0" | yarp rpc /icubSim/world
echo "world mk box 0.1 0.1 0.1     0.15 1.0 0.5   1 0 0" | yarp rpc /icubSim/world
echo "set pos 0 -35" | yarp rpc "/icubSim/head/rpc:i"

eval './imageReader' & > /dev/tty   #creates a reading port "/image/in" (receiving from cam/left) and an output port "/targetFound" to send the position of the target found (pure blue object)

eval './looker' & > /dev/tty # this program receives from "imageReader" the information about the location of the blue object (via port "/eyePort") and it moves the gaze of iCub accordingly

sleep 5

yarp connect /icubSim/cam/left /image/in
yarp connect /targetFound /eyePort


read
