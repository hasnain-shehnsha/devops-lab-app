pipeline {
    agent any

    stages {
        stage('Clone Repository') {
            steps {
                git branch: 'main', url: 'https://github.com/YOUR_USERNAME/devops-lab-app.git'
            }
        }
        stage('Build Docker Image') {
            steps {
                sh 'docker build -t devops-lab-app:latest .'
            }
        }
        stage('Run Docker Container') {
            steps {
                sh 'docker stop devops-lab-app || true'
                sh 'docker rm devops-lab-app || true'
                sh 'docker run -d --name devops-lab-app -p 5000:5000 devops-lab-app:latest'
            }
        }
        stage('Verify') {
            steps {
                sh 'sleep 3 && curl http://localhost:5000'
            }
        }
    }
}
