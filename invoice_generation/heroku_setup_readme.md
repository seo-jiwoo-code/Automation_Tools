To set up a Heroku project to run a Python script periodically to perform a task, you'll need to follow these steps:

1. Prerequisites:
   - Make sure you have the Heroku CLI installed on your computer.
   - Create a Heroku account if you don't have one.

2. Prepare your Python script:
   - Write the Python script that performs the desired task. For example, let's say your script is named `task_script.py`, and it's located in the root directory of your project.

    Create a runtime.txt file:
    In your project's root directory, create a file named runtime.txt.
    Open the runtime.txt file and specify the Python version you want to use. For Python 3, use the following format:

    python-3.9.7

3. Create a Procfile:
   - A Procfile is required to tell Heroku how to run your Python script. Create a file named `Procfile` (without any file extension) in your project's root directory.
   - Open the `Procfile` and add the following line:

   ```
   clock: python task_script.py
   ```

   The `clock` process type is used to run scheduled jobs on Heroku.

4. Set up a virtual environment (optional but recommended):
   - Create a virtual environment for your Python dependencies. This will ensure that the correct versions of the required packages are used.
   - In your project's root directory, run:

   ```bash
   python -m venv venv
   ```

   Activate the virtual environment:

   - On Windows:

   ```bash
   venv\Scripts\activate
   ```

   - On macOS and Linux:

   ```bash
   source venv/bin/activate
   ```

5. Install required dependencies:
   - Make sure your `task_script.py` has a list of dependencies in a `requirements.txt` file in the root directory. Add the necessary Python packages and their versions to this file.
   - Install the dependencies by running:

   ```bash
   pip install -r requirements.txt
   ```

6. Initialize a Git repository:
   - If you haven't already, initialize a Git repository in your project directory:

   ```bash
   git init
   ```

7. Commit your changes:
   - Add your files to the Git repository and commit your changes:

   ```bash
   git add .
   git commit -m "Initial commit"
   ```

8. Login to Heroku:
   - Open your terminal or command prompt and log in to Heroku using the following command:

   ```bash
   heroku login
   ```

   This will open a browser window prompting you to log in to your Heroku account.

9. Create a new Heroku app:
   - Create a new Heroku app with the following command:

   ```bash
   heroku create your-app-name
   ```

   Replace `your-app-name` with a unique name for your Heroku app. This name will be used in the app's URL.

10. Push your code to Heroku:
    - Deploy your code to Heroku using Git:

    ```bash
    git push heroku master
    ```

11. Set up the scheduler add-on:
    - Heroku provides an add-on called Heroku Scheduler, which allows you to schedule tasks to run at specified intervals. To add this to your app, run:

    ```bash
    heroku addons:create scheduler:standard
    ```

12. Configure the scheduled task:
    - Open the Scheduler dashboard:

    ```bash
    heroku addons:open scheduler
    ```

    - Click on "Add Job" and set the frequency and next run time for your Python script. Use the following command to run your script:

    ```
    python task_script.py
    ```

    Save the job configuration.

13. Verify and monitor:
    - Your Python script will now be scheduled to run at the specified intervals on Heroku. You can monitor the logs and check the output of your script using the Heroku CLI:

    ```bash
    heroku logs --tail
    ```

That's it! Your Python script will now run periodically as scheduled on Heroku using the Heroku Scheduler add-on. Remember to test your application and ensure everything works as expected before relying on it for critical tasks.