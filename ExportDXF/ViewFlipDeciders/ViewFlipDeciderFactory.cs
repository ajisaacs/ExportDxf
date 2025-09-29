using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExportDXF.ViewFlipDeciders
{
    /// <summary>
    /// Factory for discovering and creating IViewFlipDecider implementations.
    /// </summary>
    public static class ViewFlipDeciderFactory
    {
        private static readonly object _lock = new object();
        private static List<IViewFlipDecider> _cachedDeciders;

        /// <summary>
        /// Gets all available IViewFlipDecider implementations from loaded assemblies.
        /// Results are cached for performance.
        /// </summary>
        /// <returns>An enumerable collection of IViewFlipDecider instances.</returns>
        public static IEnumerable<IViewFlipDecider> GetAvailableDeciders()
        {
            if (_cachedDeciders != null)
                return _cachedDeciders;

            lock (_lock)
            {
                if (_cachedDeciders != null)
                    return _cachedDeciders;

                _cachedDeciders = DiscoverDeciders().ToList();
                return _cachedDeciders;
            }
        }

        /// <summary>
        /// Gets a specific view flip decider by name.
        /// </summary>
        /// <param name="name">The name of the decider to retrieve.</param>
        /// <returns>The matching decider, or null if not found.</returns>
        public static IViewFlipDecider GetDeciderByName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;

            return GetAvailableDeciders()
                .FirstOrDefault(d => string.Equals(d.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Gets the default (Automatic) view flip decider, or the first available if Automatic doesn't exist.
        /// </summary>
        /// <returns>The default decider.</returns>
        public static IViewFlipDecider GetDefaultDecider()
        {
            var automatic = GetDeciderByName("Automatic");
            if (automatic != null)
                return automatic;

            return GetAvailableDeciders().FirstOrDefault();
        }

        /// <summary>
        /// Clears the cached deciders, forcing rediscovery on next call.
        /// Useful for testing or if assemblies are loaded dynamically.
        /// </summary>
        public static void ClearCache()
        {
            lock (_lock)
            {
                _cachedDeciders = null;
            }
        }

        /// <summary>
        /// Discovers all IViewFlipDecider implementations in loaded assemblies.
        /// </summary>
        private static IEnumerable<IViewFlipDecider> DiscoverDeciders()
        {
            var deciderType = typeof(IViewFlipDecider);
            var discoveredTypes = new List<Type>();

            try
            {
                // Search all loaded assemblies
                var assemblies = AppDomain.CurrentDomain.GetAssemblies();

                foreach (var assembly in assemblies)
                {
                    try
                    {
                        var types = assembly.GetTypes()
                            .Where(t => deciderType.IsAssignableFrom(t) &&
                                       t.IsClass &&
                                       !t.IsAbstract &&
                                       HasParameterlessConstructor(t));

                        discoveredTypes.AddRange(types);
                    }
                    catch (ReflectionTypeLoadException)
                    {
                        // Skip assemblies that can't be loaded
                        continue;
                    }
                }
            }
            catch (Exception)
            {
                // If discovery fails entirely, return empty collection
                yield break;
            }

            // Create instances of discovered types
            foreach (var type in discoveredTypes)
            {
                IViewFlipDecider instance = null;

                try
                {
                    instance = (IViewFlipDecider)Activator.CreateInstance(type);
                }
                catch (Exception)
                {
                    // Skip types that can't be instantiated
                    continue;
                }

                if (instance != null)
                {
                    yield return instance;
                }
            }
        }

        /// <summary>
        /// Checks if a type has a public parameterless constructor.
        /// </summary>
        private static bool HasParameterlessConstructor(Type type)
        {
            return type.GetConstructor(
                BindingFlags.Public | BindingFlags.Instance,
                null,
                Type.EmptyTypes,
                null) != null;
        }
    }
}