node {
  stage('Checkout'){
    git 'https://github.com/nandarao/PyramidPoc.git'
  }
  stage('complit-package'){
    dif mvnHome = tool name: 'MAVEN_HOME', type: 'maven'
    sh '${mvnHome}/bin/mvn package' 
  }
}
