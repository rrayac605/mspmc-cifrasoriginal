# Introduction 
Microservisio encargado de obtener las cifras originales para PMC 

#Variables de entorno
============== AMBIENTE DE QA ==============

Para que este microservicio pueda ser ejecutado deberán de existir las siguietnes variables de entorno en el SO annfitrión donde se ejecute el jar
- portMSCifras 8033
- authenticationdatabaseMongo PMCQA01
- usrMongoCifras pmcmsconsulta
- pwMongoCifras pmcmsconsulta0
- databaseMongo PMCQA01
- portMongo 27017
- hostMongo 10.100.8.78
- fileLogCifras /weblogic/pmc/logs/mspmc-cifrasoriginal.log

============== AMBIENTE DE UAT ==============

Para que este microservicio pueda ser ejecutado deberán de existir las siguietnes variables de entorno en el SO annfitrión donde se ejecute el jar
- portMSCifras 8033
- authenticationdatabaseMongo PMCUAT01
- usrMongoCifras pmcmsconsulta
- pwMongoCifras pmcmsconsulta0
- databaseMongo PMCUAT01
- portMongo 27017
- hostMongo 10.100.8.80
- fileLogCifras /weblogic/pmc/logs/mspmc-cifrasoriginal.log

============== AMBIENTE DE PROD ==============

Para que este microservicio pueda ser ejecutado deberán de existir las siguietnes variables de entorno en el SO annfitrión donde se ejecute el jar
- portMSCifras 8033
- authenticationdatabaseMongo PENDIENTE
- usrMongoCifras pmcmsconsulta
- pwMongoCifras pmcmsconsulta0
- databaseMongo PENDIENTE
- portMongo 27017
- hostMongo PENDIENTE
- fileLogCifras /weblogic/pmc/logs/mspmc-cifrasoriginal.log

# Getting Started
TODO: Guide users through getting your code up and running on their own system. In this section you can talk about:
1.	Installation process
2.	Software dependencies
3.	Latest releases
4.	API references

# Build and Test
TODO: Describe and show how to build your code and run the tests. 

# Contribute
TODO: Explain how other users and developers can contribute to make your code better. 

If you want to learn more about creating good readme files then refer the following [guidelines](https://docs.microsoft.com/en-us/azure/devops/repos/git/create-a-readme?view=azure-devops). You can also seek inspiration from the below readme files:
- [ASP.NET Core](https://github.com/aspnet/Home)
- [Visual Studio Code](https://github.com/Microsoft/vscode)
- [Chakra Core](https://github.com/Microsoft/ChakraCore)