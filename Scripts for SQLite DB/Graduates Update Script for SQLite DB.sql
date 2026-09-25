--Author:  Kevin Thomas, Director of Admissions and Records
--Date:    23-DEC-2023
--Purpose: This script inserts the grad date  in the database.  Historically, this wasn't 
--         imported from Banner into Diplomas on Demand.  


update graduates
set grad_date = 'Tenth of December, 2022'
where term = 'Fall 2022';

