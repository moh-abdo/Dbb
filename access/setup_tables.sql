-- Create tables and example queries for Car Rental Company (Jet SQL for Access 2016+)
-- Run these in Create -> Query Design -> SQL View -> Run

-- Customers
CREATE TABLE Customers (
  CustomerID COUNTER PRIMARY KEY,
  FirstName TEXT(50),
  LastName TEXT(50),
  Email TEXT(100),
  Phone TEXT(20),
  Address TEXT(255)
);

-- Cars
CREATE TABLE Cars (
  CarID COUNTER PRIMARY KEY,
  Make TEXT(50),
  Model TEXT(50),
  Year INTEGER,
  RegNumber TEXT(20),
  DailyRate CURRENCY,
  Status TEXT(20)   -- e.g. "Available", "Rented", "Maintenance"
);

-- Rentals
CREATE TABLE Rentals (
  RentalID COUNTER PRIMARY KEY,
  CustomerID LONG,
  CarID LONG,
  StartDate DATETIME,
  EndDate DATETIME,
  ActualReturnDate DATETIME,
  DailyRate CURRENCY,
  ExtraFees CURRENCY DEFAULT 0,
  LateFees CURRENCY DEFAULT 0,
  TotalAmount CURRENCY,
  Returned YESNO DEFAULT FALSE,
  Notes TEXT(255)
);

-- Payments
CREATE TABLE Payments (
  PaymentID COUNTER PRIMARY KEY,
  RentalID LONG,
  CustomerID LONG,
  PaymentDate DATETIME,
  Amount CURRENCY,
  PaymentMethod TEXT(50),
  Notes TEXT(255)
);

-- Employees
CREATE TABLE Employees (
  EmployeeID COUNTER PRIMARY KEY,
  FirstName TEXT(50),
  LastName TEXT(50),
  Role TEXT(50),
  Email TEXT(100),
  Phone TEXT(20)
);

-- Example queries

-- Available cars for a specific period (replace [DesiredStart] and [DesiredEnd] or reference form controls)
-- Returns cars that are not rented (no overlapping active rental) and marked Available
SELECT *
FROM Cars
WHERE Cars.CarID NOT IN (
  SELECT CarID FROM Rentals
  WHERE (([DesiredStart] <= EndDate) AND ([DesiredEnd] >= StartDate)) AND Returned = False
)
AND Status = 'Available';

-- Customer rental history (prompt for CustomerID)
SELECT r.RentalID, r.CustomerID, c.Make, c.Model, r.StartDate, r.EndDate, r.TotalAmount, r.Returned
FROM Rentals AS r
INNER JOIN Cars AS c ON r.CarID = c.CarID
WHERE r.CustomerID = [EnterCustomerID];

-- Total payments per customer
SELECT p.CustomerID, Sum(p.Amount) AS TotalPaid
FROM Payments AS p
GROUP BY p.CustomerID;

-- Late returns (overdue today and not marked returned)
SELECT r.RentalID, r.CustomerID, r.CarID, r.EndDate
FROM Rentals AS r
WHERE r.Returned = False AND r.EndDate < Date();

-- Daily income (for today)
SELECT DateValue([PaymentDate]) AS PaymentDay, Sum(Amount) AS DailyIncome
FROM Payments
WHERE DateValue([PaymentDate]) = Date()
GROUP BY DateValue([PaymentDate]);
