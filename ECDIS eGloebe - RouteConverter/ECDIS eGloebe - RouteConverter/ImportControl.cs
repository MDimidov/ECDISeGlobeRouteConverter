﻿using ECDIS_eGloebe___RouteConverter.DTOs;
using ECDIS_eGloebe___RouteConverter.Utilities;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using static ECDIS_eGloebe___RouteConverter.Common.Common;

namespace ECDIS_eGloebe___RouteConverter
{
	public partial class ImportControl : UserControl
	{
		ImportRouteDto routeDto = new ImportRouteDto();

		public ImportControl()
		{
			InitializeComponent();
		}

		private void btnImportRoute_Click(object sender, EventArgs e)
		{
			ImportFromXml();
		}

		private void ImportFromXml()
		{
			string xmlStirng = File.ReadAllText(GetProjectDirectory());
			xmlStirng = xmlStirng.Replace("http://www.sevencs.com/eglobe/route", "");
			routeDto = new XmlHelper()
				.Deserialize<ImportRouteDto>(xmlStirng, "");

			CalculateAllDistancesBetweenWp();
			CalculateDistanceToGo();
		}

		private static string GetProjectDirectory()
		{
			string relativePath = @"../../../";

			return relativePath + "route.rte";
		}

		private (double, double) GetDistanceFromLastWp(
			ImportPositionDto lastPsn,
			ImportPositionDto currentPsn)
		{
			double RSH = (currentPsn.Latitude - lastPsn.Latitude) * 60 + 0.0000000000001;

			double RD = (currentPsn.Longtitude - lastPsn.Longtitude) * 60;

			double latM = lastPsn.Latitude + (RSH / 120);

			double OTSH = RD * Math.Cos(latM * Math.PI / 180);

			double course = Math.Abs(Math.Atan(OTSH / RSH) * 180 / Math.PI);

			if (RSH < 0 && RD >= 0)
			{
				course = 180 - course;
			}
			else if (RSH < 0 && RD < 0)
			{
				course = 180 + course;
			}
			else if (RSH >= 0 && RD < 0)
			{
				course = 360 - course;
			}

			double distance = Math.Round(RSH / Math.Cos(course * Math.PI / 180), 1);

			course = Math.Round(course, 1);

			return (course, distance);
		}

		private void CalculateAllDistancesBetweenWp()
		{
			for (int i = 1; i < routeDto.Waipoints.Length; i++)
			{
				(routeDto.Waipoints[i].Course, routeDto.Waipoints[i].DistanceFromLastWp) =
					GetDistanceFromLastWp(
					routeDto.Waipoints[i - 1].Position,
					routeDto.Waipoints[i].Position);
			}
		}

		private void CalculateDistanceToGo()
		{
			double totalDistance = routeDto
				.Waipoints
				.Sum(w => w.DistanceFromLastWp);

			foreach (var wp in routeDto.Waipoints)
			{
				totalDistance -= wp.DistanceFromLastWp;
				wp.DistanceToGo = Math.Round(totalDistance, 1);
			}
		}

		private void btnExportRoute_Click(object sender, EventArgs e)
		{
			ExportToWordDoc();
		}

		//private void ExportRouteToFile()
		//{
		//	StringBuilder sb = new StringBuilder();

		//	int cnter = 0;
		//	foreach (var wp in routeDto.Waipoints)
		//	{
		//		string text = $"{++cnter}. lat = {wp.Position.LatDegrees}° {wp.Position.LatMinutes.ToString("f1")}' {wp.Position.LatDir} / Long = {wp.Position.LongDegrees}° {wp.Position.LongMinutes.ToString("f1")}' {wp.Position.LongDir} / Course = {wp.Course}° / Distance = {wp.DistanceFromLastWp} n.mi. / Distance to go = {wp.DistanceToGo} / Notes = NIL";
		//		sb.AppendLine(text);
		//	}

		//	File.WriteAllText("../../../routeExport.txt", sb.ToString());
		//}

	}
}
