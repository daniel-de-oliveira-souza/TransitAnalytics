# Transit Analytics

## Overview
Transit Analytics is an innovative project that analyzes and visualizes transit data from various agencies. This project involves processing transit data provided in Excel files, executing data manipulation to derive key insights, and presenting these insights through summary statistics and visuals on a Hugo-based static website.

## Features
- **Data Integration:** Automates the consolidation of transit data from multiple agencies, provided in Excel format.
- **Advanced Data Processing:** Employs Python-based scripts for efficient data cleaning, processing, and analysis.
- **Dynamic Visualization:** Generates compelling visuals and summary statistics, clearly interpreting the data.
- **Hugo Website Display:** Presents processed data and insights on a static website built with Hugo, ensuring accessibility and ease of understanding.

## Getting Started

### Prerequisites
- Git
- Hugo
- Python 3.x
- Python Libraries: pandas, seaborn, matplotlib

### Installation and Setup
1. **Clone the repository:**
git clone https://github.com/dsouza14/transit-analytics.git
cd transit-analytics


2. **Python Environment Setup:**
- Create and activate a virtual environment (recommended):
  ```
  python -m venv venv
  source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
  ```
- Install required libraries:
  ```
  pip install -r requirements.txt
  ```

3. **Processing the Data:**
- Place your Excel files in the `data/` directory.
- Run the script to process the data:
  ```
  python main_logistics.py
  ```

4. **Building the Hugo Site:**
- Make sure Hugo is installed on your system.
- Navigate to the website directory and start the Hugo server:
  ```
  hugo server
  ```

### How to Use
- Add new Excel files to the `data/` folder.
- Execute `main_logistics.py` to update the data processing.
- Rebuild the Hugo site to display the updated analysis and visualizations.

## Contributing
Your contributions are welcome! Here are some ways you can contribute:
- **Reporting Issues:** Use the GitHub issues tab to report bugs or suggest enhancements.
- **Pull Requests:** Feel free to fork the repository and submit pull requests with your proposed changes.

## License
This project is licensed under the [MIT License](LICENSE).

## Contact Information
- **Name:** [Daniel De Oliveira Souza]
- **Email:** [danieldeoliveira1993@gmail.com]
- **Project Link:** [https://github.com/dsouza14/transit-analytics](https://github.com/dsouza14/transit-analytics)
