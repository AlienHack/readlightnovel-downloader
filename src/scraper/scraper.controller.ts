import {Controller, Get, Param, Post} from '@nestjs/common';
import {ScraperService} from "./scraper.service";

@Controller('scraper')
export class ScraperController {
    constructor(private scraperService: ScraperService) {
    }

    @Get(':novelUrl')
    async download(@Param('novelUrl') novelUrl: string){
        let buffer = new Buffer(novelUrl, 'base64');
        return this.scraperService.scrapingNovel(buffer.toString('utf8'))
    }
}
