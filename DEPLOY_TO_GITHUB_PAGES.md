# Make the add-in work without keeping Command Prompt open

The current Excel test add-in loads from `https://localhost:3000`. Closing its webpack window stops that local site. This project now includes a GitHub Pages publishing setup so the add-in can load from an always-online HTTPS address instead.

## One-time setup

1. Sign in to GitHub or create a free GitHub account.
2. Create a **public** repository named `filtered-column-finder`. GitHub Pages on the free plan requires a public repository. The repository contains the add-in's source code; it does not contain your Excel workbooks.
3. Upload this project's files to that repository, including the `.github/workflows/deploy-pages.yml` file. Do not upload the `node_modules` or `dist` folders.
4. In the repository on GitHub, open **Settings → Pages**. Under **Build and deployment**, set **Source** to **GitHub Actions**.
5. Open the **Actions** tab and wait for **Build and publish Excel add-in** to finish successfully.
6. Your permanent add-in website will be `https://YOUR-GITHUB-NAME.github.io/filtered-column-finder/`.

## Connect Excel to the hosted version

1. In GitHub, open the published website address and add `/manifest.xml` to the end. Download that `manifest.xml` file to your computer.
2. In Excel desktop, open **Insert → Add-ins → My Add-ins** (or **Home → Add-ins → More Add-ins**, depending on your ribbon).
3. Choose **Upload My Add-in**, select the downloaded `manifest.xml`, and add it.
4. Open the add-in's **Filtered Column Finder** ribbon group and choose **Go to Last Filtered Column**.

After this, Excel loads the add-in from GitHub Pages; you can close the local webpack window. The existing developer-installed copy may still show an error. Remove that old test copy from **My Add-ins**, then use the hosted manifest copy.

## Important

- This avoids the Microsoft Marketplace publisher account. The add-in is for your own Excel installation; this does not publish it in Microsoft's store or install it automatically into every device/account.
- The add-in source files on a public GitHub Pages site are visible to anyone. The add-in does not upload your workbook data to GitHub; it uses the Excel JavaScript API to inspect the active sheet and move the selection.
- GitHub Pages is online hosting, so access to the hosting account is needed for future updates. Its availability depends on GitHub and your internet connection.
