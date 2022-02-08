package main

import (
	"log"
	"strings"

	"golang.org/x/sys/windows/registry"
)

const (
	OUTLOOKSETTINGS = `SOFTWARE\Microsoft\Office`
	OUTLOOKPROFILES = `SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles`
	SETTINGS        = `9375CFF0413111d3B88A00104B2A6676`
)

func main() {
	setSmtpAddressInProfile()
	setSmtpAddressInSettings()
}

func setSmtpAddress(s registry.Key) {
	var smtpKey string
	smtpTemp, valType, _ := s.GetBinaryValue("SMTP Server")
	if valType == 1 {
		smtpKey, _, _ = s.GetStringValue("SMTP Server")
	} else {
		smtpKey = strings.Replace(string(smtpTemp), "\x00", "", -1)
	}
	if smtpKey == "exchange.svetlana-k.ru" {
		err := s.SetDWordValue("SMTP Port", 587)
		if err != nil {
			log.Println(err)
		}
	}
}

func getSubkeys(k registry.Key) (subKeys []string, err error) {
	subKeys, err = k.ReadSubKeyNames(-1)
	if err != nil {
		log.Println(err)
	}
	return subKeys, err
}

func setSmtpAddressInSettings() {
	k, err := registry.OpenKey(
		registry.CURRENT_USER,
		OUTLOOKSETTINGS,
		registry.ENUMERATE_SUB_KEYS)

	if err != nil {
		log.Println(err)
	}
	defer k.Close()

	versions, err := getSubkeys(k)
	if err != nil {
		log.Println(err)

	}

	for _, version := range versions {
		versionSettings, err := registry.OpenKey(
			registry.CURRENT_USER,
			OUTLOOKSETTINGS+"\\"+version+"\\"+"Outlook\\Profiles\\Outlook"+"\\"+SETTINGS,
			registry.READ)
		if err != nil {
			log.Println(err)
		}
		defer k.Close()

		sk, err := getSubkeys(versionSettings)
		if err != nil {
			log.Println(err)
		}

		for _, v := range sk {
			settings, err := registry.OpenKey(
				registry.CURRENT_USER,
				OUTLOOKSETTINGS+"\\"+version+"\\"+"Outlook\\Profiles\\Outlook"+"\\"+SETTINGS+"\\"+v,
				registry.ALL_ACCESS)
			if err != nil {
				log.Println(err)
			}
			defer settings.Close()
			setSmtpAddress(settings)

		}
	}
}

func setSmtpAddressInProfile() {
	k, err := registry.OpenKey(
		registry.CURRENT_USER,
		OUTLOOKPROFILES,
		registry.ENUMERATE_SUB_KEYS)

	if err != nil {
		log.Println(err)
	}
	defer k.Close()

	profiles, err := getSubkeys(k)
	if err != nil {
		log.Println(err)
	}

	for _, profileName := range profiles {
		profileSettings, err := registry.OpenKey(
			registry.CURRENT_USER,
			OUTLOOKPROFILES+"\\"+profileName+"\\"+SETTINGS,
			registry.READ)
		if err != nil {
			log.Println(err)
		}
		defer k.Close()

		sk, err := getSubkeys(profileSettings)
		if err != nil {
			log.Println(err)
		}

		for _, v := range sk {
			settings, err := registry.OpenKey(
				registry.CURRENT_USER,
				OUTLOOKPROFILES+"\\"+profileName+"\\"+SETTINGS+"\\"+v,
				registry.ALL_ACCESS)
			if err != nil {
				log.Println(err)
			}
			defer settings.Close()
			setSmtpAddress(settings)

		}
	}
}
