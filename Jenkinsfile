node {
  stage('Checkout'){
    git 'https://github.com/nandarao/PyramidPoc.git'
  }
  stage('complit-package'){
    def mvnHome = tool name: 'Maven', type: 'maven'
    sh "${mvnHome}/bin/mvn package" 
  }
}
