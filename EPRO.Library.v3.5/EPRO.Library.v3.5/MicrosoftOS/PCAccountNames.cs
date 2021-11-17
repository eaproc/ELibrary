using System;

namespace ELibrary.Standard.MicrosoftOS
{
    public class PCAccountNames
    {
        public enum AccountNames : byte
        {
            ADMINISTRATOR,
            ADMINISTRATORS,
            EVERYONE,
            GUEST,
            GUESTS,
            LOCAL__SERVICE,
            NETWORK,
            NETWORK__SERVICE,
            SERVICE,
            SYSTEM,
            USERS
        }

        private OSCulture ___InUseCulture;

        public OSCulture InUseCulture
        {
            get
            {
                return ___InUseCulture;
            }
        }

        public PCAccountNames(OSCulture pCulture)
        {
            ___InUseCulture = pCulture;
        }

        public string getName(AccountNames pAccount)
        {
            switch (InUseCulture.ClassCultureType)
            {
                case OSCulture.Cultures.ENGLISH___________USA:
                case OSCulture.Cultures.ENGLISH___________UK:
                    {
                        switch (pAccount)
                        {
                            case AccountNames.ADMINISTRATOR:
                                {
                                    return "ADMINISTRATOR";
                                }

                            case AccountNames.ADMINISTRATORS:
                                {
                                    return "ADMINISTRATOR";
                                }

                            case AccountNames.EVERYONE:
                                {
                                    return "EVERYONE";
                                }

                            case AccountNames.GUEST:
                                {
                                    return "GUEST";
                                }

                            case AccountNames.GUESTS:
                                {
                                    return "GUESTS";
                                }

                            case AccountNames.LOCAL__SERVICE:
                                {
                                    return "LOCAL SERVICE";
                                }

                            case AccountNames.NETWORK:
                                {
                                    return "NETWORK";
                                }

                            case AccountNames.NETWORK__SERVICE:
                                {
                                    return "NETWORK SERVICE";
                                }

                            case AccountNames.SERVICE:
                                {
                                    return "SERVICE";
                                }

                            case AccountNames.SYSTEM:
                                {
                                    return "SYSTEM";
                                }

                            case AccountNames.USERS:
                                {
                                    return "USERS";
                                }

                            default:
                                {
                                    throw new Exception("This Account Name is NOT supported. " + pAccount.ToString());
                                    break;
                                }
                        }

                        break;
                    }

                case OSCulture.Cultures.FRENCH_________FRANCE:
                    {
                        switch (pAccount)
                        {
                            case AccountNames.ADMINISTRATOR:
                                {
                                    return "Administrateur";
                                }

                            case AccountNames.ADMINISTRATORS:
                                {
                                    return "Administrateurs";
                                }

                            case AccountNames.EVERYONE:
                                {
                                    return "Tout le monde";
                                }

                            case AccountNames.GUEST:
                                {
                                    return "Invité";
                                }

                            case AccountNames.GUESTS:
                                {
                                    return "Invités";
                                }

                            case AccountNames.LOCAL__SERVICE:
                                {
                                    return "SERVICE LOCAL";
                                }

                            case AccountNames.NETWORK:
                                {
                                    return "RESEAU";
                                }

                            case AccountNames.NETWORK__SERVICE:
                                {
                                    return "SERVICE RÉSEAU";
                                }

                            case AccountNames.SERVICE:
                                {
                                    return "SERVICE";
                                }

                            case AccountNames.SYSTEM:
                                {
                                    return "SYSTÈME";
                                }

                            case AccountNames.USERS:
                                {
                                    return "Utilisateurs";
                                }

                            default:
                                {
                                    throw new Exception("This Account Name is NOT supported. " + pAccount.ToString());
                                    break;
                                }
                        }

                        break;
                    }

                default:
                    {
                        throw new Exception("This culture is not Supported Yet. " + Environment.NewLine + ___InUseCulture.getClassCultureSummary());
                        break;
                    }
            }
        }
    }
}