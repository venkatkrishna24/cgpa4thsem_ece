const gradeMap = { "O": 10, "A+": 9, "A": 8, "B+": 7, "B": 6, "C": 5, "U": 0 };
let subjectCount = 9;

const subjectsData = [
  { name: "C23411 COMMUNICATION SYSTEM", credit: 3 },
  { name: "EC23412 ELECTROMAGNETIC FIELD", credit: 4 },
  { name: "EC23413 LINEAR INTEGRATED CIRCUIT", credit: 3 },
  { name: "EC23431 DIGITAL SIGNAL PROCESSING", credit: 4 },
  { name: "EC23432 NETWORK AND SECURITY", credit: 3 },
  { name: "MA23412 RANDOM PROCESS & LINEAR ALGEBRA", credit: 4 },
  { name: "EC23421 COMMUNICATION SYSTEM LAB", credit: 1 },
  { name: "EC23422 LINEAR INTEGRATED CIRCUIT LAB", credit: 1 },
  { name: "EC23IC2 CYBER SECURITY", credit: 1 }
];

function createSubjectField(subject, index) {
  return `
    <div class="subject" data-default="true" data-index="${index}">
      <label>${subject.name} (${subject.credit} Credits)</label>
      <select id="sub${index}" data-credit="${subject.credit}" required>
        <option value="">Select Grade</option>
        <option>O</option><option>A+</option><option>A</option><option>B+</option><option>B</option><option>C</option><option>U</option>
      </select>
    </div>
  `;
}

function initializeSubjects() {
  const container = document.getElementById('subjectsContainer');
  subjectsData.forEach((subject, i) => {
    container.insertAdjacentHTML('beforeend', createSubjectField(subject, i+1));
  });
}
initializeSubjects();

function calculateCGPA(showResult = true) {
  const subjects = document.querySelectorAll('#subjectsContainer select');
  let totalCredits = 0, weightedSum = 0;
  for (let sub of subjects) {
    const credit = parseInt(sub.getAttribute('data-credit'));
    const grade = sub.value;
    if (grade === "") {
      alert("Please select grade for all subjects.");
      return null;
    }
    weightedSum += gradeMap[grade] * credit;
    totalCredits += credit;
  }
  const cgpa = (weightedSum / totalCredits).toFixed(2);
  if (showResult) {
    const resultDiv = document.getElementById("result");
    resultDiv.innerHTML = `<strong>CGPA: ${cgpa}</strong>`;
    resultDiv.style.display = 'block';
    setTimeout(() => { resultDiv.style.display = 'none'; }, 4000);
  }
  return cgpa;
}

function addExtraSubject() {
  subjectCount++;
  const container = document.getElementById('subjectsContainer');
  const div = document.createElement('div');
  div.classList.add('subject');
  div.innerHTML = `
    <label>Extra Subject ${subjectCount}</label>
    <input type="text" class="extra-name" placeholder="Subject Name" required>
    <input type="number" class="extra-credit" placeholder="Credits" min="1" value="1" required>
    <select id="sub${subjectCount}" data-credit="1" required>
      <option value="">Select Grade</option>
      <option>O</option><option>A+</option><option>A</option><option>B+</option><option>B</option><option>C</option><option>U</option>
    </select>
  `;
  container.appendChild(div);

  const creditInput = div.querySelector('.extra-credit');
  const select = div.querySelector('select');
  creditInput.addEventListener('input', () => {
    select.setAttribute('data-credit', creditInput.value);
  });
  div.scrollIntoView({ behavior: "smooth" });
}

function downloadExcel() {
  const name = document.getElementById("name").value.trim();
  const roll = document.getElementById("roll").value.trim();
  if (!name || !roll) return alert("Please enter your Name and Roll Number first.");

  const cgpa = calculateCGPA(false);
  if (!cgpa) return;

  const wb = XLSX.utils.book_new();
  const wsData = [
    ["Name", name],
    ["Roll Number", roll],
    [],
    ["Subject", "Credits", "Grade"]
  ];

  const allSubjects = document.querySelectorAll('#subjectsContainer .subject');
  allSubjects.forEach(div => {
    let subjectName = div.dataset.default === "true"
      ? div.querySelector('label').innerText.split(' (')[0]
      : div.querySelector('.extra-name').value.trim() || 'Unnamed Subject';

    const credit = div.querySelector('select')?.getAttribute('data-credit') || 0;
    const grade = div.querySelector('select')?.value || '';
    wsData.push([subjectName, credit, grade]);
  });

  wsData.push([], ["Final CGPA", cgpa]);

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(wb, ws, "CGPA");
  XLSX.writeFile(wb, `${roll}_CGPA.xlsx`);
}
