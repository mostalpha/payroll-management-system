import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

const initialForm = {
  id: '',
  employeeName: '',
  employeeId: '',
  department: '',
  month: '',
  basicSalary: '',
  overtimeHours: '',
  overtimeRate: '',
  allowances: '',
  deductions: ''
};

const emptyForm = () => ({ ...initialForm });

const formatCurrency = (value) => {
  const number = Number(value || 0);
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD'
  }).format(number);
};

function App() {
  const [formData, setFormData] = useState(emptyForm());
  const [payrollList, setPayrollList] = useState([]);
  const [isEditing, setIsEditing] = useState(false);

  useEffect(() => {
    const saved = localStorage.getItem('payroll-data');
    if (saved) {
      setPayrollList(JSON.parse(saved));
    }
  }, []);

  useEffect(() => {
    localStorage.setItem('payroll-data', JSON.stringify(payrollList));
  }, [payrollList]);

  const totalGross = useMemo(
    () =>
      payrollList.reduce((sum, employee) => {
        const gross =
          Number(employee.basicSalary || 0) +
          Number(employee.allowances || 0) +
          Number(employee.overtimeHours || 0) * Number(employee.overtimeRate || 0);
        return sum + gross;
      }, 0),
    [payrollList]
  );

  const totalNet = useMemo(
    () =>
      payrollList.reduce((sum, employee) => {
        const overtimePay = Number(employee.overtimeHours || 0) * Number(employee.overtimeRate || 0);
        const net =
          Number(employee.basicSalary || 0) +
          Number(employee.allowances || 0) +
          overtimePay -
          Number(employee.deductions || 0);
        return sum + net;
      }, 0),
    [payrollList]
  );

  const handleChange = (event) => {
    const { name, value } = event.target;
    setFormData((prev) => ({ ...prev, [name]: value }));
  };

  const calculateNetPay = (employee) => {
    const overtimePay = Number(employee.overtimeHours || 0) * Number(employee.overtimeRate || 0);
    return (
      Number(employee.basicSalary || 0) +
      Number(employee.allowances || 0) +
      overtimePay -
      Number(employee.deductions || 0)
    );
  };

  const handleSubmit = (event) => {
    event.preventDefault();

    const cleanedData = {
      ...formData,
      basicSalary: Number(formData.basicSalary || 0),
      overtimeHours: Number(formData.overtimeHours || 0),
      overtimeRate: Number(formData.overtimeRate || 0),
      allowances: Number(formData.allowances || 0),
      deductions: Number(formData.deductions || 0)
    };

    if (!cleanedData.employeeName || !cleanedData.employeeId || !cleanedData.month) {
      alert('Please fill in employee name, employee ID, and payroll month.');
      return;
    }

    if (isEditing && cleanedData.id) {
      setPayrollList((prev) =>
        prev.map((employee) => (employee.id === cleanedData.id ? { ...cleanedData } : employee))
      );
    } else {
      const newEntry = {
        ...cleanedData,
        id: Date.now().toString()
      };
      setPayrollList((prev) => [newEntry, ...prev]);
    }

    setFormData(emptyForm());
    setIsEditing(false);
  };

  const handleEdit = (employee) => {
    setFormData(employee);
    setIsEditing(true);
  };

  const handleDelete = (id) => {
    setPayrollList((prev) => prev.filter((employee) => employee.id !== id));
    if (isEditing) {
      setFormData(emptyForm());
      setIsEditing(false);
    }
  };

  const exportToExcel = () => {
    if (payrollList.length === 0) {
      alert('No payroll records to export.');
      return;
    }

    const exportRows = payrollList.map((employee) => ({
      'Employee Name': employee.employeeName,
      'Employee ID': employee.employeeId,
      'Department': employee.department,
      'Month': employee.month,
      'Basic Salary': employee.basicSalary,
      'Overtime Hours': employee.overtimeHours,
      'Overtime Rate': employee.overtimeRate,
      'Allowances': employee.allowances,
      'Deductions': employee.deductions,
      'Net Pay': calculateNetPay(employee)
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Payroll');
    XLSX.writeFile(workbook, 'payroll-export.xlsx');
  };

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">Payroll Dashboard</p>
          <h1>Payroll Management System</h1>
        </div>
        <button className="primary-btn" onClick={exportToExcel}>
          Download Excel
        </button>
      </header>

      <section className="summary-grid">
        <div className="summary-card">
          <span>Total Employees</span>
          <strong>{payrollList.length}</strong>
        </div>
        <div className="summary-card">
          <span>Total Gross</span>
          <strong>{formatCurrency(totalGross)}</strong>
        </div>
        <div className="summary-card">
          <span>Total Net</span>
          <strong>{formatCurrency(totalNet)}</strong>
        </div>
      </section>

      <main className="content-grid">
        <section className="panel form-panel">
          <h2>{isEditing ? 'Edit Payroll Entry' : 'Add Payroll Entry'}</h2>

          <form onSubmit={handleSubmit} className="payroll-form">
            <div className="field-row two-col">
              <label>
                Employee Name
                <input
                  type="text"
                  name="employeeName"
                  value={formData.employeeName}
                  onChange={handleChange}
                  placeholder="e.g. Ada Johnson"
                />
              </label>

              <label>
                Employee ID
                <input
                  type="text"
                  name="employeeId"
                  value={formData.employeeId}
                  onChange={handleChange}
                  placeholder="e.g. EMP-001"
                />
              </label>
            </div>

            <div className="field-row two-col">
              <label>
                Department
                <input
                  type="text"
                  name="department"
                  value={formData.department}
                  onChange={handleChange}
                  placeholder="e.g. Finance"
                />
              </label>

              <label>
                Payroll Month
                <input
                  type="month"
                  name="month"
                  value={formData.month}
                  onChange={handleChange}
                />
              </label>
            </div>

            <div className="field-row two-col">
              <label>
                Basic Salary
                <input
                  type="number"
                  name="basicSalary"
                  value={formData.basicSalary}
                  onChange={handleChange}
                  min="0"
                  step="0.01"
                />
              </label>

              <label>
                Overtime Hours
                <input
                  type="number"
                  name="overtimeHours"
                  value={formData.overtimeHours}
                  onChange={handleChange}
                  min="0"
                  step="0.5"
                />
              </label>
            </div>

            <div className="field-row two-col">
              <label>
                Overtime Rate
                <input
                  type="number"
                  name="overtimeRate"
                  value={formData.overtimeRate}
                  onChange={handleChange}
                  min="0"
                  step="0.01"
                />
              </label>

              <label>
                Allowances
                <input
                  type="number"
                  name="allowances"
                  value={formData.allowances}
                  onChange={handleChange}
                  min="0"
                  step="0.01"
                />
              </label>
            </div>

            <div className="field-row">
              <label>
                Deductions
                <input
                  type="number"
                  name="deductions"
                  value={formData.deductions}
                  onChange={handleChange}
                  min="0"
                  step="0.01"
                />
              </label>
            </div>

            <div className="form-actions">
              <button type="submit" className="primary-btn">
                {isEditing ? 'Update Entry' : 'Save Entry'}
              </button>
              {isEditing && (
                <button
                  type="button"
                  className="secondary-btn"
                  onClick={() => {
                    setFormData(emptyForm());
                    setIsEditing(false);
                  }}
                >
                  Cancel
                </button>
              )}
            </div>
          </form>
        </section>

        <section className="panel table-panel">
          <h2>Payroll Records</h2>

          {payrollList.length === 0 ? (
            <div className="empty-state">No payroll entries yet.</div>
          ) : (
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Name</th>
                    <th>ID</th>
                    <th>Month</th>
                    <th>Net Pay</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {payrollList.map((employee) => (
                    <tr key={employee.id}>
                      <td>{employee.employeeName}</td>
                      <td>{employee.employeeId}</td>
                      <td>{employee.month}</td>
                      <td>{formatCurrency(calculateNetPay(employee))}</td>
                      <td className="actions">
                        <button type="button" className="mini-btn edit" onClick={() => handleEdit(employee)}>
                          Edit
                        </button>
                        <button type="button" className="mini-btn delete" onClick={() => handleDelete(employee.id)}>
                          Delete
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </section>
      </main>
    </div>
  );
}

export default App;
