/* Copyright (C) 2010, Manuel Meitinger
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 2 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;

namespace Aufbauwerk.Tools.GroupPolicyInstaller
{
    public partial class MainForm : Form
    {
        private IEnumerator<Task> tasks;
        private Task task;

        public MainForm()
        {
            InitializeComponent();
        }

        private void StartNextProcess()
        {
            // reset everything
            task = null;
            logoPictureBox.Image = null;
            productLabel.Text = string.Empty;
            progressBar.Value = 100;
            progressBar.Style = ProgressBarStyle.Marquee;
            actionLabel.Text = string.Empty;

            // check if there are more tasks
            if (tasks.MoveNext())
            {
                // initialize the task
                task = tasks.Current;
                logoPictureBox.Image = task.Image;
                productLabel.Text = task.Name;
                ThreadPool.QueueUserWorkItem(RunTaskAsync);
            }
            else
            {
                // release the list enumerator and exit the application
                tasks.Dispose();
                Application.Exit();
            }
        }

        private void RunTaskAsync(object state)
        {
            // run a task (this is called in another thread!)
            task.Run((action, progress) => Invoke(new TaskProgressChanged(ProgressChanged), action, progress));
            Invoke(new MethodInvoker(StartNextProcess));
        }

        private void ProgressChanged(string action, int progress)
        {
            // set the action text and progress bar
            actionLabel.Text = action;
            if (progress != -1)
            {
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = progress;
            }
            else
                progressBar.Style = ProgressBarStyle.Marquee;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // get the tasks and start with the first one
            tasks = Program.Tasks.GetEnumerator();
            StartNextProcess();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // prevent the form from being closed by the user
            e.Cancel = e.CloseReason == CloseReason.UserClosing;
        }
    }
}
