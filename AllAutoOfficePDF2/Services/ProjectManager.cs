using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using AllAutoOfficePDF2.Models;

namespace AllAutoOfficePDF2.Services
{
    /// <summary>
    /// �v���W�F�N�g�Ǘ��T�[�r�X
    /// </summary>
    public class ProjectManager
    {
        private static readonly string ProjectsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "AllAutoOfficePDF2",
            "projects.json"
        );

        /// <summary>
        /// �v���W�F�N�g��ǂݍ���
        /// </summary>
        /// <returns>�v���W�F�N�g���X�g</returns>
        public static List<ProjectData> LoadProjects()
        {
            try
            {
                if (File.Exists(ProjectsFilePath))
                {
                    var json = File.ReadAllText(ProjectsFilePath);
                    var projects = JsonSerializer.Deserialize<List<ProjectData>>(json) ?? new List<ProjectData>();
                    return projects;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"�v���W�F�N�g�̓ǂݍ��݂Ɏ��s���܂���: {ex.Message}", "�G���[",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
            return new List<ProjectData>();
        }

        /// <summary>
        /// �v���W�F�N�g��ۑ�
        /// </summary>
        /// <param name="projects">�v���W�F�N�g���X�g</param>
        public static void SaveProjects(List<ProjectData> projects)
        {
            try
            {
                var directory = Path.GetDirectoryName(ProjectsFilePath);
                if (!Directory.Exists(directory))
                    Directory.CreateDirectory(directory!);

                var json = JsonSerializer.Serialize(projects, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
                File.WriteAllText(ProjectsFilePath, json);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"�v���W�F�N�g�̕ۑ��Ɏ��s���܂���: {ex.Message}", "�G���[",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }
    }
}