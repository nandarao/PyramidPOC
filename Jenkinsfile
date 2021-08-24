pipeline {
agent any

tools {
maven "3.8.2"
}

stages {
stage("build "){
steps {
sh "mvn -version"
sh "mvn clean install"
}
}
}
post  {
always  {
cleanWs()
}
}
}