pipeline {

    agent any
  
    tools {
        maven 'MAVEN_HOME' 
    }
  
    stages {
       stage('Checkout'){
         steps {
           git 'https://github.com/nandarao/PyramidPoc.git'
         }
    }
        stage('Compile stage') {
            steps {
                bat "mvn clean compile" 
        }
    }

         stage('testing stage') {
             steps {
                bat "mvn test"
        }
    }

          stage('deployment stage') {
              steps {
                bat "mvn deploy"
        }
    }

  }

}
