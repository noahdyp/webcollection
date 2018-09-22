#!/bin/bash
pwd=`pwd`
JAVA_HOME=$pwd/jre
`chmod 777 $JAVA_HOME/bin/java`
$JAVA_HOME/bin/java -jar webcollection.jar $pwd/result/;