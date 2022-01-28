package main

import (
	"log"
	"strings"

	"golang.org/x/sys/windows/registry"
)

const (
	OUTLOOKPROFILES = `SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles`
	SETTINGS        = `9375CFF0413111d3B88A00104B2A6676`
)

func main() {
	k, err := registry.OpenKey(
		registry.CURRENT_USER,
		OUTLOOKPROFILES,
		registry.ENUMERATE_SUB_KEYS)

	if err != nil {
		log.Fatal(err)
	}
	defer k.Close()

	profiles, err := getSubkeys(k)
	if err != nil {
		log.Fatal(err)
	}

	for _, profileName := range profiles {
		profileSettings, err := registry.OpenKey(
			registry.CURRENT_USER,
			OUTLOOKPROFILES+"\\"+profileName+"\\"+SETTINGS,
			registry.READ)
		if err != nil {
			log.Fatal(err)
		}
		defer k.Close()
		setSmtpAddress(profileSettings, profileName)

	}

}

func setSmtpAddress(ps registry.Key, pn string) {
	sk, err := getSubkeys(ps)
	if err != nil {
		log.Fatal(err)
	}

	for _, v := range sk {
		settings, err := registry.OpenKey(
			registry.CURRENT_USER,
			OUTLOOKPROFILES+"\\"+pn+"\\"+SETTINGS+"\\"+v,
			registry.ALL_ACCESS)
		if err != nil {
			log.Fatal(err)
		}
		defer settings.Close()

		smtpTemp, _, _ := settings.GetBinaryValue("SMTP Server")

		smtpKey := strings.Replace(string(smtpTemp), "\x00", "", -1)

		if smtpKey == "exchange.svetlana-k.ru" {
			err = settings.SetDWordValue("SMTP Port", 465)
			if err != nil {
				log.Fatal(err)
			}
		}
	}
}

func getSubkeys(k registry.Key) (subKeys []string, err error) {
	subKeys, err = k.ReadSubKeyNames(-1)
	if err != nil {
		log.Fatal(err)
	}
	return subKeys, err
}
