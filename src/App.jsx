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

const calculateGrossPay = (employee) => {
  return (
    Number(employee.basicSalary || 0) +
    Number(employee.allowances || 0) +
    Number(employee.overtimeHours || 0) * Number(employee.overtimeRate || 0)
  );
};

const calculateNetPay = (employee) => {
  return calculateGrossPay(employee) - Number(employee.deductions || 0);
};

function App() {
  const [formData, setFormData] = useState(emptyForm());
  const [payrollList, setPayrollList] = useState([]);
  const [isEditing, setIsEditing] = useState(false);
  const [errors, setErrors] = useState({});
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedMonth, setSelectedMonth] = useState('');

  useEffect(() => {
    const saved = localStorage.getItem('payroll-data');
    if (saved) {
      setPayrollList(JSON.parse(saved));
    }
  }, []);

  useEffect(() => {
    localStorage.setItem('payroll-data', JSON.stringify(payrollList));
  }, [payrollList]);

  const filteredPayrollList = useMemo(() => {
    return payrollList.filter((employee) => {
      const matchesSearch =
        !searchTerm ||
        employee.employeeName.toLowerCase().includes(searchTerm.toLowerCase()) ||
        employee.employeeId.toLowerCase().includes(searchTerm.toLowerCase());

      const matchesMonth = !selectedMonth || employee.month === selectedMonth;
      return matchesSearch && matchesMonth;
    });
  }, [payrollList, searchTerm, selectedMonth]);

  const totalGross = useMemo(
    () => payrollList.reduce((sum, employee) => sum + calculateGrossPay(employee), 0),
    [payrollList]
  );

  const totalDeductions = useMemo(
    () => payrollList.reduce((sum, employee) => sum + Number(employee.deductions || 0), 0),
    [payrollList]
  );

  const totalNet = useMemo(
    () => payrollList.reduce((sum, employee) => sum + calculateNetPay(employee), 0),
    [payrollList]
  );

  const averageNet = useMemo(() => {
    if (payrollList.length === 0) return 0;
    return totalNet / payrollList.length;
  }, [payrollList, totalNet]);

  const handleChange = (event) => {
    const { name, value } = event.target;
    setFormData((prev) => ({ ...prev, [name]: value }));
    setErrors((prev) => ({ ...prev, [name]: '' }));
  };

  const validateForm = () => {
    const nextErrors = {};

    if (!formData.employeeName.trim()) nextErrors.employeeName = 'Employee name is required.';
    if (!formData.employeeId.trim()) nextErrors.employeeId = 'Employee ID is required.';
    if (!formData.month) nextErrors.month = 'Payroll month is required.';

    if (Number(formData.basicSalary || 0) < 0) nextErrors.basicSalary = 'Basic salary cannot be negative.';
    if (Number(formData.overtimeHours || 0) < 0) nextErrors.overtimeHours = 'Overtime hours cannot be negative.';
    if (Number(formData.overtimeRate || 0) < 0) nextErrors.overtimeRate = 'Overtime rate cannot be negative.';
    if (Number(formData.allowances || 0) < 0) nextErrors.allowances = 'Allowances cannot be negative.';
    if (Number(formData.deductions || 0) < 0) nextErrors.deductions = 'Deductions cannot be negative.';

    if (Number(formData.basicSalary || 0) === 0 && !formData.employeeName.trim()) {
      nextErrors.basicSalary = 'Basic salary is required.';
    }

    setErrors(nextErrors);
    return Object.keys(nextErrors).length === 0;
  };

  const handleSubmit = (event) => {
    event.preventDefault();

    if (!validateForm()) {
      return;
    }

    const cleanedData = {
      ...formData,
      employeeName: formData.employeeName.trim(),
      employeeId: formData.employeeId.trim(),
      department: formData.department.trim(),
      basicSalary: Number(formData.basicSalary || 0),
      overtimeHours: Number(formData.overtimeHours || 0),
      overtimeRate: Number(formData.overtimeRate || 0),
      allowances: Number(formData.allowances || 0),
      deductions: Number(formData.deductions || 0)
    };

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
    setErrors({});
  };

  const handleEdit = (employee) => {
    setFormData(employee);
    setIsEditing(true);
    setErrors({});
  };

  const handleDelete = (id) => {
    const record = payrollList.find((employee) => employee.id === id);
    const confirmDelete = window.confirm(
      `Delete payroll for ${record?.employeeName || 'this employee'}?`
    );

    if (!confirmDelete) return;

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
      'Gross Pay': calculateGrossPay(employee),
      'Net Pay': calculateNetPay(employee)
    }));

    const totalsRow = {
      'Employee Name': 'TOTAL',
      'Employee ID': '',
      'Department': '',
      'Month': '',
      'Basic Salary': totalGross,
      'Overtime Hours': '',
      'Overtime Rate': '',
      'Allowances': '',
      'Deductions': totalDeductions,
      'Gross Pay': totalGross,
      'Net Pay': totalNet
    };

    const worksheet = XLSX.utils.json_to_sheet([...exportRows, totalsRow]);
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
          <span>Total Deductions</span>
          <strong>{formatCurrency(totalDeductions)}</strong>
        </div>
        <div className="summary-card">
          <span>Total Net</span>
          <strong>{formatCurrency(totalNet)}</strong>
        </div>
        <div className="summary-card">
          <span>Average Net</span>
          <strong>{formatCurrency(averageNet)}</strong>
        </div>
      </section>

      <main className="content-grid">
        <section className="panel form-panel">
          <h2>{isEditing ? 'Edit Payroll Entry' : 'Add Payroll Entry'}</h2>

          <form onSubmit={handleSubmit} className="payroll-form" noValidate>
            <div className="field-row two-col">
              <label>
                Employee Name
                <input
                  type="text"
                  name="employeeName"
                  value={formData.employeeName}
                  onChange={handleChange}
                  placeholder="e.g. Ada Johnson"
                  className={errors.employeeName ? 'input-error' : ''}
                />
                {errors.employeeName && <span className="error-text">{errors.employeeName}</span>}
              </label>

              <label>
                Employee ID
                <input
                  type="text"
                  name="employeeId"
                  value={formData.employeeId}
                  onChange={handleChange}
                  placeholder="e.g. EMP-001"
                  className={errors.employeeId ? 'input-error' : ''}
                />
                {errors.employeeId && <span className="error-text">{errors.employeeId}</span>}
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
                  className={errors.month ? 'input-error' : ''}
                />
                {errors.month && <span className="error-text">{errors.month}</span>}
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
                  className={errors.basicSalary ? 'input-error' : ''}
                />
                {errors.basicSalary && <span className="error-text">{errors.basicSalary}</span>}
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
                  className={errors.overtimeHours ? 'input-error' : ''}
                />
                {errors.overtimeHours && <span className="error-text">{errors.overtimeHours}</span>}
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
                  className={errors.overtimeRate ? 'input-error' : ''}
                />
                {errors.overtimeRate && <span className="error-text">{errors.overtimeRate}</span>}
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
                  className={errors.allowances ? 'input-error' : ''}
                />
                {errors.allowances && <span className="error-text">{errors.allowances}</span>}
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
                  className={errors.deductions ? 'input-error' : ''}
                />
                {errors.deductions && <span className="error-text">{errors.deductions}</span>}
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
                    setErrors({});
                  }}
                >
                  Cancel
                </button>
              )}
            </div>
          </form>
        </section>

        <section className="panel table-panel">
          <div className="table-header">
            <h2>Payroll Records</h2>
            <div className="filters">
              <input
                type="text"
                placeholder="Search by name or ID"
                value={searchTerm}
                onChange={(event) => setSearchTerm(event.target.value)}
              />
              <input
                type="month"
                value={selectedMonth}
                onChange={(event) => setSelectedMonth(event.target.value)}
              />
            </div>
          </div>

          {filteredPayrollList.length === 0 ? (
            <div className="empty-state">No payroll entries match your search.</div>
          ) : (
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Name</th>
                    <th>ID</th>
                    <th>Month</th>
                    <th>Gross</th>
                    <th>Net Pay</th>
                    <th>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredPayrollList.map((employee) => (
                    <tr key={employee.id}>
                      <td>{employee.employeeName}</td>
                      <td>{employee.employeeId}</td>
                      <td>{employee.month}</td>
                      <td>{formatCurrency(calculateGrossPay(employee))}</td>
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
