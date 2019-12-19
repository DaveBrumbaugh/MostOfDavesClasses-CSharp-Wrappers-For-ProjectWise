using System.Windows.Forms;

namespace PolyhedraCE
{
    /// <summary>
    /// Keyins Class
    /// </summary>
    public sealed class Keyins
    {
        public static void OpenForm(string unparsed)
        {
            MessageBox.Show("This seems to work.", "PolyhedraCE");
        }

        private static void Polyhedron(string unparsed)
        {
            CreatePolyhedronX64.InstallNewInstance(unparsed);
        }

        public static void SnubCube(string unparsed)
        {
            CreatePolyhedronX64.InstallNewInstance("SnubCube");
        }
        public static void TruncatedIcosahedron(string unparsed)
        {
            CreatePolyhedronX64.InstallNewInstance("TruncatedIcosahedron");
        }
        public static void Icosahedron(string unparsed)
        {
            CreatePolyhedronX64.InstallNewInstance("Icosahedron");
        }
        public static void Dodecahedron(string unparsed)
        {
            CreatePolyhedronX64.InstallNewInstance("Dodecahedron");
        }
        public static void TruncatedDodecahedron(string unparsed)
        {
            CreatePolyhedronX64.InstallNewInstance("TruncatedDodecahedron");
        }
        public static void SnubDodecahedron(string unparsed)
        {
            CreatePolyhedronX64.InstallNewInstance("SnubDodecahedron");
        }
    }
}