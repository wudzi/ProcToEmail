# ProcToEmail

Windows Service to take an input of a stored procedure (MSSQL or ASE) and send email as HTML to recipients. 

Replacement Variables are dynamically set by column names e.g. a column "forename" will replace "[$forename$]" in the HTML template

An example of Options.xml

<?xml version="1.0"?>
<Process Name="_ProcToEmail">

	<Queues>
		<Queue Name="NewEmployee">
			<LogDirectory>C:\NewEmployee\LOG</LogDirectory>
			<ArchiveFolder>C:\NewEmployee\ARCHIVE</ArchiveFolder>
			<ErrorFolder>C:\NewEmployee\ERROR</ErrorFolder>
			<DatabaseType>Sybase</DatabaseType>
			<ConnectionString>Data Source=server1;Port=5000;Database=dbname;Uid=sa;Pwd=password;</ConnectionString>
			<StoredProcedure>sp_test</StoredProcedure>
			<InputParameter>
				<Parameter Name="@DateFrom">12 oct 2020</Parameter>
				<Parameter Name="@DateTo">13 oct 2020</Parameter>
			</InputParameter>
			<HTMLTemplate>C:\NewEmployee\Template\NewEmployee.html</HTMLTemplate>
			<EmailField>email_address</EmailField>
			<IDField>employee_code</IDField>
			<AutoParameters>1</AutoParameters>
			<ParameterMapping>
				<DataMap DBField="Forename">[$Forename$]</DataMap>
				<DataMap DBField="Surname">[$Surname$]</DataMap>
			</ParameterMapping>
			<SMTPServer>send.test.co.uk</SMTPServer>
			<FromAddress>noreply@test.co.uk</FromAddress>
			<FromName>test Distribution Services</FromName>
			<!--<BCCAddress>jon@me.com</BCCAddress>-->
			<Subject>Your Employment Contract</Subject>
			<DebugEmail>j@test.co.uk</DebugEmail>
			<DebugMode>1</DebugMode>
		</Queue>

	</Queues>

</Process>
