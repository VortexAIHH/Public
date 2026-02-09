// SharePoint Permissions Inventory - RESPECTS 5000 ITEM THRESHOLD
// Fetches ALL items in batches < 5000, checks permissions client-side
// No filters that trigger HTTP 500
// Run this in browser console while logged into SharePoint

(async function () {
    console.log('🔍 SharePoint Permissions Inventory (5K Batch Limit) - Starting...');
    console.log('⚠️  Respecting SharePoint 5000 item threshold...\n');

    const permissions = [];
    const siteUrl = _spPageContextInfo.webAbsoluteUrl;
    const siteName = _spPageContextInfo.webTitle;

    console.log(`📍 Site: ${siteName}`);
    console.log(`📍 URL: ${siteUrl}\n`);

    // Helper to make REST API calls with retry logic for throttling
    async function getJSON(url, retries = 5) {
        for (let attempt = 0; attempt < retries; attempt++) {
            try {
                const response = await fetch(url, {
                    headers: { 'Accept': 'application/json;odata=verbose' },
                    credentials: 'include'
                });

                // Handle throttling - wait and retry with MINUTE delays
                if (response.status === 406 || response.status === 429) {
                    // 1min, 3min, 6min, 12min, 15min (max)
                    const waitMinutes = attempt === 0 ? 1 : Math.min(15, 3 * Math.pow(2, attempt - 1));
                    const waitTime = waitMinutes * 60 * 1000;
                    console.log(`      ⏳ Throttled (HTTP ${response.status}), waiting ${waitMinutes} minute(s) before retry ${attempt + 1}/${retries}...`);
                    await delay(waitTime);
                    continue;
                }

                if (!response.ok) throw new Error(`HTTP ${response.status}`);
                return await response.json();
            } catch (e) {
                if (attempt === retries - 1) throw e;
                const waitMinutes = attempt === 0 ? 1 : Math.min(15, 3 * Math.pow(2, attempt - 1));
                const waitTime = waitMinutes * 60 * 1000;
                console.log(`      ⚠️  Error (${e.message}), waiting ${waitMinutes} minute(s) before retry ${attempt + 1}/${retries}...`);
                await delay(waitTime);
            }
        }
        throw new Error(`Failed after ${retries} retries`);
    }

    // Helper to add delay between requests (prevents throttling)
    function delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    // Helper to get group members
    async function getGroupMembers(groupId) {
        try {
            const members = await getJSON(`${siteUrl}/_api/web/sitegroups/GetById(${groupId})/users`);
            return members.d.results.map(u => u.Title || u.LoginName).join('; ');
        } catch (e) {
            return 'Unable to retrieve members';
        }
    }

    // Helper to get permissions for any object
    async function getPermissions(objectUrl, scope, location, itemName = '') {
        try {
            const perms = await getJSON(`${objectUrl}/roleassignments?$expand=RoleDefinitionBindings,Member`);
            const results = [];

            for (const assignment of perms.d.results) {
                const member = assignment.Member;
                const roles = assignment.RoleDefinitionBindings.results.map(r => r.Name).join(', ');

                // Get group members if this is a SharePoint group
                let groupMembers = '';
                if (member.PrincipalType === 8) { // SharePoint Group
                    console.log(`         Fetching members of group: ${member.Title}`);
                    groupMembers = await getGroupMembers(member.Id);
                }

                results.push({
                    Scope: scope,
                    Location: location,
                    ItemName: itemName,
                    HasUniquePermissions: 'Yes',
                    PrincipalType: getPrincipalTypeName(member.PrincipalType),
                    PrincipalName: member.Title || '',
                    PrincipalLoginName: member.LoginName || '',
                    GroupMembers: groupMembers,
                    Roles: roles
                });
            }
            return results;
        } catch (e) {
            console.warn(`   ⚠️  Could not get permissions: ${e.message}`);
            return [];
        }
    }

    function getPrincipalTypeName(type) {
        const types = { 1: 'User', 4: 'SharePoint Group', 8: 'Security Group' };
        return types[type] || `Type ${type}`;
    }

    try {
        // 1. SITE-LEVEL PERMISSIONS
        console.log('🔐 Step 1/3: Getting site-level permissions...');
        const sitePerms = await getPermissions(
            `${siteUrl}/_api/web`,
            'Site',
            siteUrl,
            siteName
        );
        permissions.push(...sitePerms);
        console.log(`   ✅ Found ${sitePerms.length} site-level permission entries\n`);

        // 2. GET ALL LISTS/LIBRARIES
        console.log('📚 Step 2/3: Getting all lists and libraries...');
        const allLists = await getJSON(
            `${siteUrl}/_api/web/lists?` +
            `$select=Title,Id,HasUniqueRoleAssignments,ItemCount,BaseType,Hidden&` +
            `$filter=Hidden eq false`
        );

        // Filter out system lists
        const lists = allLists.d.results.filter(list => {
            const systemLists = [
                'appdata', 'appfiles', 'Composed Looks', 'Converted Forms', 'Form Templates',
                'List Template Gallery', 'Master Page Gallery', 'Site Assets', 'Site Pages',
                'Solution Gallery', 'Style Library', 'Theme Gallery', 'User Information List',
                'Web Part Gallery', 'Content and Structure Reports', 'Reporting Metadata',
                'Reporting Templates', 'Content type publishing error log', 'Cache Profiles',
                'Project Policy Item List', 'Workflow History', 'Workflow Tasks',
                'Access Requests', 'TaxonomyHiddenList', 'Site Collection Documents',
                'Site Collection Images', 'Preservation Hold Library', 'Pages',
                'Relationships List', 'Reusable Content', 'Quick Deploy Items',
                'Long Running Operation Status', 'Maintenance Log Library', 'Hub Sites',
                'Site Script Catalog', 'Site Design Catalog'
            ];
            const lowerTitle = list.Title.toLowerCase();
            if (systemLists.some(sys => lowerTitle.includes(sys.toLowerCase()))) return false;
            return list.BaseType === 0 || list.BaseType === 1;
        });

        console.log(`   ✅ Found ${lists.length} user lists/libraries\n`);

        // 3. LIST-LEVEL AND ITEM-LEVEL PERMISSIONS
        console.log('📋 Step 3/3: Scanning all lists and items...');

        let totalItemsScanned = 0;
        let totalUniquePerms = 0;

        for (let i = 0; i < lists.length; i++) {
            const list = lists[i];
            console.log(`\n   [${i + 1}/${lists.length}] ${list.Title} (${list.ItemCount.toLocaleString()} items)`);

            // List-level permissions
            if (list.HasUniqueRoleAssignments) {
                const listPerms = await getPermissions(
                    `${siteUrl}/_api/web/lists(guid'${list.Id}')`,
                    list.BaseType === 1 ? 'Library' : 'List',
                    `${siteUrl} > ${list.Title}`,
                    list.Title
                );
                permissions.push(...listPerms);
                console.log(`      ✅ List has ${listPerms.length} unique permissions`);
                totalUniquePerms += listPerms.length;
            } else {
                console.log(`      (inherits from site)`);
            }

            // Item-level: Fetch in batches of 1000 (well under 5000 threshold)
            if (list.ItemCount > 0) {
                console.log(`      Scanning items in batches of 1000...`);

                try {
                    let skip = 0;
                    const batchSize = 1000; // Safe batch size under 5K threshold
                    let itemsWithUniquePerms = 0;

                    while (skip < list.ItemCount) {
                        // NO FILTER - just fetch items in batches
                        const batch = await getJSON(
                            `${siteUrl}/_api/web/lists(guid'${list.Id}')/items?` +
                            `$select=Id,FileLeafRef,FileRef,HasUniqueRoleAssignments,FSObjType&` +
                            `$top=${batchSize}&$skip=${skip}`
                        );

                        if (batch.d.results.length === 0) break;

                        // Check each item CLIENT-SIDE for unique permissions
                        for (const item of batch.d.results) {
                            if (item.HasUniqueRoleAssignments) {
                                const itemName = item.FileLeafRef || `Item ${item.Id}`;
                                const itemPath = item.FileRef || '';
                                const scope = item.FSObjType === 1 ? 'Folder' : (list.BaseType === 1 ? 'File' : 'Item');

                                const itemPerms = await getPermissions(
                                    `${siteUrl}/_api/web/lists(guid'${list.Id}')/items(${item.Id})`,
                                    scope,
                                    `${siteUrl} > ${list.Title} > ${itemPath}`,
                                    itemName
                                );
                                permissions.push(...itemPerms);
                                itemsWithUniquePerms++;
                                totalUniquePerms += itemPerms.length;

                                // Small delay to prevent throttling
                                await delay(50);
                            }
                        }

                        skip += batchSize;
                        totalItemsScanned += batch.d.results.length;

                        // Delay between batches to prevent throttling
                        await delay(100);

                        if (skip % 5000 === 0) {
                            console.log(`      Scanned ${skip.toLocaleString()} items, found ${itemsWithUniquePerms} with unique perms...`);
                        }
                    }

                    console.log(`      ✅ Scanned ${totalItemsScanned.toLocaleString()} items, found ${itemsWithUniquePerms} with unique permissions`);

                } catch (e) {
                    console.error(`      ❌ Error scanning items: ${e.message}`);
                }
            }
        }

        // GENERATE EXCEL FILE (HTML Table Method)
        console.log('\n💾 Generating Excel file...');

        // Create HTML table for Excel
        let html = `
            <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
            <head>
                <meta charset="UTF-8">
                <!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>
                <x:Name>Permissions Report</x:Name>
                <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
                </x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->
            </head>
            <body>
                <table border="1">
                    <thead style="background-color: #4472C4; color: white; font-weight: bold;">
                        <tr>
                            <th>Scope</th>
                            <th>Location</th>
                            <th>Item Name</th>
                            <th>Has Unique Permissions</th>
                            <th>Principal Type</th>
                            <th>Principal Name</th>
                            <th>Principal Login</th>
                            <th>Group Members</th>
                            <th>Roles</th>
                        </tr>
                    </thead>
                    <tbody>`;

        for (const perm of permissions) {
            // Color-code rows by scope
            let bgColor = '#FFFFFF';
            switch (perm.Scope) {
                case 'Site': bgColor = '#E7E6FF'; break;
                case 'List': bgColor = '#D9E1F2'; break;
                case 'Library': bgColor = '#D9E1F2'; break;
                case 'Folder': bgColor = '#FFF2CC'; break;
                case 'File': bgColor = '#E2EFDA'; break;
                case 'Item': bgColor = '#E2EFDA'; break;
            }

            html += `<tr style="background-color: ${bgColor};">
                <td>${perm.Scope}</td>
                <td>${perm.Location}</td>
                <td>${perm.ItemName}</td>
                <td style="font-weight: bold; color: #C00000;">${perm.HasUniquePermissions}</td>
                <td>${perm.PrincipalType}</td>
                <td>${perm.PrincipalName}</td>
                <td>${perm.PrincipalLoginName}</td>
                <td style="font-size: 9px;">${perm.GroupMembers || ''}</td>
                <td>${perm.Roles}</td>
            </tr>`;
        }

        html += `</tbody></table></body></html>`;

        // Download as Excel file
        const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Permissions_Report_${siteName.replace(/[^a-z0-9]/gi, '_')}_${new Date().toISOString().slice(0, 10)}.xls`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);

        // SUMMARY
        console.log('\n' + '='.repeat(80));
        console.log('📊 PERMISSIONS REPORT GENERATED');
        console.log('='.repeat(80));
        console.log(`\n   Total Permission Entries: ${permissions.length.toLocaleString()}`);
        console.log(`   Total Items Scanned: ${totalItemsScanned.toLocaleString()}`);
        console.log(`   Items with Unique Permissions: ${totalUniquePerms.toLocaleString()}`);
        console.log(`\n   Site-Level: ${permissions.filter(p => p.Scope === 'Site').length}`);
        console.log(`   List/Library-Level: ${permissions.filter(p => p.Scope === 'List' || p.Scope === 'Library').length}`);
        console.log(`   Folder-Level: ${permissions.filter(p => p.Scope === 'Folder').length}`);
        console.log(`   File/Item-Level: ${permissions.filter(p => p.Scope === 'File' || p.Scope === 'Item').length}`);
        console.log('\n✅ Excel file downloaded successfully!');
        console.log('='.repeat(80) + '\n');

    } catch (error) {
        console.error('❌ Fatal Error:', error);
        console.error(error.stack);
    }
})();
