// Initialize MSAL Configuration
const msalConfig = {
    auth: {
        clientId: "<your-client-id-goes-here>", 
        authority: "https://login.microsoftonline.com/<your-tenant-id-goes-here>", 
        redirectUri: "http://localhost:8000",
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false,
    },
};

// Create MSAL instance
let msalInstance;
try {
    msalInstance = new msal.PublicClientApplication(msalConfig);
    console.log("MSAL Instance initialized successfully.");
} catch (error) {
    console.error("Error initializing MSAL instance:", error);
}

// Pagination Variables
let allGroups = [];
let groupData = [];
let allGroupNames = [];
let allGroupOwners = [];
let currentPage = 1;
const itemsPerPage = 10;
let currentData = [];

// Login
async function login() {
    if (!msalInstance) {
        console.error("MSAL instance is not initialized!");
        document.getElementById("output").innerText = "Login failed: MSAL instance is not initialized.";
        return;
    }

    try {
        const loginResponse = await msalInstance.loginPopup({
            scopes: ["Group.Read.All", "User.Read"],
        });
        console.log("Login successful:", loginResponse);
        msalInstance.setActiveAccount(loginResponse.account);
        alert('Login Successful');
        fetchGroups();
    } catch (error) {
        console.error("Login failed:", error);
        alert('Login Failed');
    }
}

// Logout
function logout() {
    if (!msalInstance) {
        console.error("MSAL instance is not initialized!");
        return;
    }

    msalInstance.logoutPopup();

    // Clear all data
    allGroups = [];
    groupData = [];
    currentData = [];
    allGroupNames = [];
    allGroupOwners = [];
    currentPage = 1;

    // Clear dropdowns
    document.getElementById("groupNameDropdown").innerHTML = '<option value="">Select Group Name</option>';
    document.getElementById("groupOwnerDropdown").innerHTML = '<option value="">Select Group Owner</option>';
    document.getElementById("groupVisibilityDropdown").value = '';

    // Clear UI
    document.querySelector(".results-container").innerHTML = "";
    document.getElementById("pagination").innerHTML = "";
}


// Fetch Groups (Using the exact function you provided)
async function fetchGroups() {
    try {
        groupData = [];
        let nextLink = `/groups?$select=id,displayName,visibility`;

        while (nextLink) {
            const response = await callGraphApi(nextLink);

            if (response.value) {
                console.log(response.value)
                allGroups = allGroups.concat(response.value);

                const groupDetails = await Promise.all(
                    response.value.map(async (group) => {
                        let membersCount = "N/A";
                        let ownerEmails = "N/A";

                        try {
                            // Fetch member count with required header
                            const membersCountResponse = await fetch(`https://graph.microsoft.com/v1.0/groups/${group.id}/members/$count`, {
                                method: "GET",
                                headers: {
                                    Authorization: `Bearer ${await getAccessToken()}`,
                                    ConsistencyLevel: "eventual",
                                },
                            });

                            if (membersCountResponse.ok) {
                                membersCount = await membersCountResponse.text();
                                membersCount = parseInt(membersCount, 10) || 0;
                            } else {
                                console.warn(`Failed to fetch member count for group: ${group.displayName}. Status: ${membersCountResponse.status}`);
                            }
                        } catch (error) {
                            console.error(`Error fetching members count for group ${group.displayName}:`, error);
                        }

                        try {
                            // Fetch group owners
                            const ownersResponse = await callGraphApi(`/groups/${group.id}/owners?$select=userPrincipalName`);
                            if (ownersResponse.value) {
                                ownerEmails = ownersResponse.value.map((owner) => owner.userPrincipalName).join(", ") || "N/A";
                            } else {
                                console.warn(`Failed to fetch owners for group: ${group.displayName}`);
                            }
                        } catch (error) {
                            console.error(`Error fetching owners for group ${group.displayName}:`, error);
                        }

                        return {
                            groupName: group.displayName || "N/A",
                            groupVisibility: group.visibility || "N/A",
                            memberCount: membersCount,
                            groupOwner: ownerEmails,
                        };
                    })
                );

                groupData = groupData.concat(groupDetails);
                allGroupNames = [...new Set(groupData.map((group) => group.groupName))];
                allGroupOwners = [...new Set(groupData.flatMap((group) => group.groupOwner.split(", ")).filter(Boolean))];
            }

            nextLink = response["@odata.nextLink"];
        }

        currentData = groupData;
        changePage(1);
        populateDropdowns();
    } catch (error) {
        console.error("Error fetching groups:", error);
        alert("Failed to fetch groups. Please try again.");
    }
}


// Pagination Functions
function changePage(page) {
    currentPage = page;
    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    displayResults(currentData.slice(startIndex, endIndex));
    createPaginationControls(currentData.length);
}

function createPaginationControls(totalItems) {
    const paginationContainer = document.getElementById("pagination");
    const totalPages = Math.ceil(totalItems / itemsPerPage);
    paginationContainer.innerHTML = "";

    const prevButton = document.createElement("button");
    prevButton.textContent = "Prev";
    prevButton.className = "btn btn-secondary";
    prevButton.disabled = currentPage === 1;
    prevButton.onclick = () => changePage(currentPage - 1);
    paginationContainer.appendChild(prevButton);

    for (let i = 1; i <= totalPages; i++) {
        const pageButton = document.createElement("button");
        pageButton.textContent = i;
        pageButton.className = `btn btn-${i === currentPage ? "primary" : "light"}`;
        pageButton.onclick = () => changePage(i);
        paginationContainer.appendChild(pageButton);
    }

    const nextButton = document.createElement("button");
    nextButton.textContent = "Next";
    nextButton.className = "btn btn-secondary";
    nextButton.disabled = currentPage === totalPages;
    nextButton.onclick = () => changePage(currentPage + 1);
    paginationContainer.appendChild(nextButton);
}

// Populate Dropdowns
function populateDropdowns() {
    populateDropdown("groupNameDropdown", "Select Group Name", allGroupNames);
    populateDropdown("groupOwnerDropdown", "Select Group Owner", allGroupOwners);
}

function populateDropdown(dropdownId, placeholder, values) {
    const dropdown = document.getElementById(dropdownId);
    dropdown.innerHTML = `<option value="">${placeholder}</option>`;
    values.forEach((value) => {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = value;
        dropdown.appendChild(option);
    });
}


function search() {
    const groupName = document.getElementById("groupNameDropdown").value;
    const groupVisibility = document.getElementById("groupVisibilityDropdown").value;
    const groupOwner = document.getElementById("groupOwnerDropdown").value;

    const filteredGroups = groupData.filter((group) => {
        const matchesName = groupName ? group.groupName === groupName : true;
        const matchesVisibility = groupVisibility ? group.groupVisibility === groupVisibility : true;
        const matchesOwner = groupOwner ? group.groupOwner.includes(groupOwner) : true;
        return matchesName && matchesVisibility && matchesOwner;
    });

    currentData = filteredGroups;
    if (filteredGroups.length === 0) {
        alert("No matching results found.");
    }
    changePage(1);
}




function downloadReportAsCSV() {
    const headers = ["Group Name", "Group Visibility", "Members", "Group Owner"];
    const rows = currentData.map(group => [
        group.groupName || "N/A",
        group.groupVisibility || "N/A",
        group.memberCount || 0,
        group.groupOwner || "N/A"
    ]);

    const csvContent = [headers.join(","), ...rows.map(row => row.join(","))].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "GroupsReport.csv";
    a.click();
    URL.revokeObjectURL(url);
}



async function callGraphApi(endpoint, method = "GET", body = null, headers = {}) {
    const account = msalInstance.getActiveAccount();
    if (!account) throw new Error("Please log in first.");

    const tokenResponse = await msalInstance.acquireTokenSilent({
        scopes: ["Group.Read.All", "User.Read"],
        account,
    });

    const defaultHeaders = {
        Authorization: `Bearer ${tokenResponse.accessToken}`,
        "Content-Type": "application/json",
    };

    const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
        method,
        headers: { ...defaultHeaders, ...headers },
        body: body ? JSON.stringify(body) : null,
    });

    if (response.ok) {
        const contentType = response.headers.get("content-type");
        if (contentType && contentType.includes("application/json")) {
            return await response.json();
        }
        return {};
    } else {
        const errorText = await response.text();
        console.error("Graph API Error Response:", errorText);
        throw new Error(`Graph API call failed: ${response.status} ${response.statusText}`);
    }
}




async function sendReportAsMail() {
    const email = prompt("Enter your email address:");
    if (!email) return alert("Please provide an email address.");

    const headers = ["Group Name", "Group Visibility", "Members", "Group Owner"];

    const emailContent = currentData.map((group) =>
        `<tr>
            <td>${group.groupName || "N/A"}</td>
            <td>${group.groupVisibility || "N/A"}</td>
            <td>${group.memberCount || 0}</td>
            <td>${group.groupOwner || "N/A"}</td>
        </tr>`
    ).join("");

    const emailBody = `<table border="1">
        <thead>
            <tr>${headers.map(header => `<th>${header}</th>`).join("")}</tr>
        </thead>
        <tbody>${emailContent}</tbody>
    </table>`;

    const message = {
        message: {
            subject: "Groups Report",
            body: { contentType: "HTML", content: emailBody },
            toRecipients: [{ emailAddress: { address: email } }],
        },
    };

    try {
        await callGraphApi("/me/sendMail", "POST", message);
        alert("Report sent!");
    } catch (error) {
        console.error("Error sending report:", error);
        alert("Failed to send the report.");
    }
}





function displayResults(data) {
    const resultsContainer = document.querySelector(".results-container");

    // Clear existing content
    resultsContainer.innerHTML = "";

    // Create table dynamically
    const table = document.createElement("table");
    table.className = "table table-striped";

    // Create table header
    const thead = document.createElement("thead");
    thead.className = "outputHeader"; // Add class name for compatibility
    thead.innerHTML = `
        <tr>
            <th>Group Name</th>
            <th>Group Visibility</th>
            <th>Members</th>
            <th>Group Owner</th>
        </tr>
    `;
    table.appendChild(thead);

    // Create table body
    const tbody = document.createElement("tbody");
    tbody.className = "outputBody"; // Add class name for compatibility

    if (data.length > 0) {
        data.forEach((group) => {
            const row = document.createElement("tr");
            row.innerHTML = `
                <td>${group.groupName || "N/A"}</td>
                <td>${group.groupVisibility || "N/A"}</td>
                <td>${group.memberCount || 0}</td>
                <td>${group.groupOwner || "N/A"}</td>
            `;
            tbody.appendChild(row);
        });
    } else {
        // No data case
        const row = document.createElement("tr");
        row.innerHTML = `<td colspan="4" style="text-align: center;">No matching groups found.</td>`;
        tbody.appendChild(row);
    }

    table.appendChild(tbody);
    resultsContainer.appendChild(table);
}


async function getAccessToken() {
    const account = msalInstance.getActiveAccount();
    if (!account) throw new Error("Please log in first.");

    const tokenResponse = await msalInstance.acquireTokenSilent({
        scopes: ["Directory.Read.All"],
        account,
    });

    return tokenResponse.accessToken;
}






// Expose Functions
window.fetchGroups = fetchGroups;
window.search = search;
window.downloadReportAsCSV = downloadReportAsCSV;
window.sendReportAsMail = sendReportAsMail;
window.login = login;
window.logout = logout;

document.addEventListener("DOMContentLoaded", () => {
    document.getElementById("loginBtn").addEventListener("click", login);
    document.getElementById("logoutBtn").addEventListener("click", logout);
    document.getElementById("searchBtn").addEventListener("click", search);
    document.getElementById("downloadBtn").addEventListener("click", downloadReportAsCSV);
    document.getElementById("mailBtn").addEventListener("click", sendReportAsMail);
});
