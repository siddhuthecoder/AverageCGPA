document.addEventListener('DOMContentLoaded', () => {
    const showRanksBtn = document.getElementById('showRanksBtn');
    const searchByRankBtn = document.getElementById('searchByRankBtn');
    const searchByIdBtn = document.getElementById('searchByIdBtn');
    const rankInput = document.getElementById('rankInput');
    const idInput = document.getElementById('idInput');
    const resultsDiv = document.getElementById('results');
    const userImage = document.getElementById('userImage');
    const ranksTable = document.getElementById('ranksTable');
    const ranksTableBody = document.getElementById('ranksTableBody');

    let studentData = [];

    // Function to fetch and parse Excel file
    async function loadExcelFile() {
        try {
            const response = await fetch('student_average_grades.xlsx');
            const arrayBuffer = await response.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet);
            studentData = rows.map(row => ({
                id: row['Student ID'],
                averageGrade: row['Average Grade']
            }));
            console.log(studentData); // Check the parsed data
        } catch (error) {
            console.error('Error loading the Excel file:', error);
        }
    }

    // Load data by default
    loadExcelFile();

    // Function to display data with rank in a table
    function displayRanks() {
        resultsDiv.innerHTML = '';
        ranksTable.style.display = 'table';
        ranksTableBody.innerHTML = '';
        studentData
            .map((student, index) => ({
                rank: index + 1,
                ...student
            }))
            .forEach(item => {
                ranksTableBody.innerHTML += `
                    <tr>
                        <td>${item.rank}</td>
                        <td>${item.id}</td>
                        <td>${item.averageGrade.toFixed(2)}</td>
                        <td><img src="https://intranet.rguktn.ac.in/SMS/usrphotos/user/${item.id}.jpg" alt="Student Image" style="width:50px; height:50px;"></td>
                    </tr>`;
            });
    }

    // Function to search by rank
    function searchByRank() {
        const rank = parseInt(rankInput.value);
        if (isNaN(rank) || rank < 1 || rank > studentData.length) {
            resultsDiv.innerHTML = '<div class="result-item">Invalid rank.</div>';
            userImage.style.display = 'none';
            return;
        }
        const student = studentData[rank - 1];
        resultsDiv.innerHTML = `<div class="result-item">Rank: ${rank}, ID: ${student.id}, Average Grade: ${student.averageGrade.toFixed(2)}</div>`;
        userImage.src = `https://intranet.rguktn.ac.in/SMS/usrphotos/user/${student.id}.jpg`;
        userImage.style.display = 'block';
    }

    // Function to search by ID
    function searchById() {
        const id = idInput.value.trim();
        const student = studentData.find(stu => stu.id === id);
        if (student) {
            const rank = studentData.indexOf(student) + 1;
            resultsDiv.innerHTML = `<div class="result-item">ID: ${id}, Rank: ${rank}, Average Grade: ${student.averageGrade.toFixed(2)}</div>`;
            userImage.src = `https://intranet.rguktn.ac.in/SMS/usrphotos/user/${id}.jpg`;
            userImage.style.display = 'block';
        } else {
            resultsDiv.innerHTML = '<div class="result-item">Student ID not found.</div>';
            userImage.style.display = 'none';
        }
    }

    // Event Listeners
    showRanksBtn.addEventListener('click', displayRanks);
    searchByRankBtn.addEventListener('click', searchByRank);
    searchByIdBtn.addEventListener('click', searchById);
});
